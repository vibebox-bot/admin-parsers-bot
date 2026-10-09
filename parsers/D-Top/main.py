import os
import json
import re
import time
import sys
from datetime import datetime
from urllib.parse import (
    urljoin,
    urlparse,
    parse_qs,
    parse_qsl,
    urlencode,
    urlunparse,
)

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Харьковская D-Top — NEW SITE")

# ==========================================================
# CONFIG
# ==========================================================

BASE = "https://www.dtopelectronics.com.ua"
LOGIN_URL = BASE + "/index.php?route=account/login"

#CATEGORY_LIMIT = None
CATEGORY_LIMIT = 2

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/D-Top")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "D-Top_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "lock.txt"
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
}

session = requests.Session()
session.headers.update(HEADERS)

# Кэш артикулов, чтобы не открывать одну страницу повторно.
SKU_CACHE = {}

# ==========================================================
# LOCK
# ==========================================================

def is_locked():
    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        if age > 3600:
            os.remove(LOCK_FILE)
            print("⚠️ Удалена устаревшая блокировка")
            return False

        return True

    except OSError:
        return False


def set_lock(state):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if state:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))
    else:
        try:
            if os.path.exists(LOCK_FILE):
                os.remove(LOCK_FILE)
        except OSError:
            pass


# ==========================================================
# STATUS
# ==========================================================

def save_status(running=False, progress=0, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path,
    }

    tmp = STATUS_PATH + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, STATUS_PATH)


# ==========================================================
# HELPERS
# ==========================================================

def clean(value):
    if not value:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def get_soup(url, attempts=3):
    for attempt in range(1, attempts + 1):
        try:
            response = session.get(url, timeout=30)

            if response.status_code == 200:
                response.encoding = response.apparent_encoding or "utf-8"
                return BeautifulSoup(response.text, "html.parser")

            print(
                f"⚠️ HTTP {response.status_code}: {url}"
            )

        except requests.RequestException as exc:
            print(
                f"⚠️ Запрос {attempt}/{attempts}: "
                f"{type(exc).__name__}"
            )

        if attempt < attempts:
            time.sleep(attempt)

    return BeautifulSoup("", "html.parser")


def absolute_url(href, base_url=BASE):
    if not href:
        return ""

    href = href.strip()

    if (
        href.startswith("#")
        or href.lower().startswith("javascript:")
        or href.lower().startswith("mailto:")
        or href.lower().startswith("tel:")
    ):
        return ""

    return urljoin(base_url, href)


def normalize_category_url(url):
    """
    Оставляем только URL категории.
    Параметры страницы и сортировки не должны
    создавать отдельные категории.
    """

    url = absolute_url(url)

    if not url:
        return ""

    parsed = urlparse(url)

    if parsed.hostname not in (
        "dtopelectronics.com.ua",
        "www.dtopelectronics.com.ua",
    ):
        return ""

    params = parse_qs(parsed.query)

    if params.get("route", [""])[0] != "product/category":
        return ""

    path = params.get("path", [""])[0]

    if not path:
        return ""

    query = urlencode({
        "route": "product/category",
        "path": path,
    })

    return urlunparse((
        "https",
        "www.dtopelectronics.com.ua",
        "/index.php",
        "",
        query,
        "",
    ))


def build_page_url(category_url, page):
    parsed = urlparse(category_url)

    params = [
        (key, value)
        for key, value in parse_qsl(
            parsed.query,
            keep_blank_values=True,
        )
        if key != "page"
    ]

    if page > 1:
        params.append(("page", str(page)))

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        urlencode(params),
        "",
    ))


# ==========================================================
# LOGIN
# ==========================================================

def login():
    print("🔐 Авторизация D-Top...")

    soup = get_soup(LOGIN_URL)

    form = soup.select_one(
        'form[action*="account/login"]'
    )

    if not form:
        print("⚠️ Не найдена форма авторизации")
        return False

    payload = {}

    for inp in form.select("input[name]"):
        name = inp.get("name")
        payload[name] = inp.get("value", "")

    payload["email"] = EMAIL
    payload["password"] = PASSWORD

    action = absolute_url(
        form.get("action"),
        LOGIN_URL,
    ) or LOGIN_URL

    try:
        response = session.post(
            action,
            data=payload,
            headers={"Referer": LOGIN_URL},
            timeout=30,
            allow_redirects=True,
        )

        response.raise_for_status()

    except requests.RequestException as exc:
        print(
            f"❌ Ошибка авторизации: {type(exc).__name__}"
        )
        return False

    # Проверяем признаки авторизованного аккаунта.
    check_soup = get_soup(BASE)

    page_text = clean(check_soup.get_text(" ")).lower()
    page_html = str(check_soup).lower()

    logged_in = any(marker in page_html for marker in (
        "route=account/logout",
        "route=account/account",
    ))

    if not logged_in:
        logged_in = any(marker in page_text for marker in (
            "выход",
            "выйти",
            "logout",
            "my account",
        ))

    if logged_in:
        print("✅ LOGIN OK")
        return True

    # Запасная проверка: на сайте цены доступны
    # зарегистрированным покупателям. Проверяем
    # страницу каталога на наличие цены.
    test_url = (
        BASE
        + "/index.php?route=product/category&path=77"
    )

    test_soup = get_soup(test_url)

    for price_el in test_soup.select(
        ".product-thumb__price"
    ):
        price = clean(price_el.get_text(" ", strip=True))

        if not price:
            price = clean(price_el.get("data-price", ""))

        if price:
            print("✅ Цены доступны после входа")
            return True

    print("❌ Авторизация не подтверждена.")
    print("❌ Цены каталога не обнаружены.")
    print("Проверь логин, пароль и доступ к сайту.")

    return False


# ==========================================================
# CATEGORIES
# ==========================================================

def get_categories():
    """
    Обходим главную страницу и все найденные категории.
    В каждой категории дополнительно ищем вложенные
    категории, пока новые ссылки не закончатся.
    """

    found = set()
    checked = set()
    queue = [BASE]

    print("📂 Сбор категорий...")

    while queue:
        current_url = queue.pop(0)

        if current_url in checked:
            continue

        checked.add(current_url)

        soup = get_soup(current_url)

        if not soup.select_one("body"):
            continue

        # Главное меню и все вложенные пункты.
        selectors = (
            "ul.menu-module__ul a[href]",
            "a.menu-module__a[href]",
            "a.menu-module__children-a[href]",
            'a[href*="route=product/category"]',
        )

        links = {}

        for selector in selectors:
            for a in soup.select(selector):
                href = normalize_category_url(
                    absolute_url(
                        a.get("href", ""),
                        current_url,
                    )
                )

                if href:
                    links[href] = href

        for category_url in links:
            if category_url not in found:
                found.add(category_url)

            if category_url not in checked:
                queue.append(category_url)

        if current_url != BASE:
            time.sleep(0.05)

    categories = sorted(found)

    print(f"📂 TOTAL CATEGORIES: {len(categories)}")

    return categories


# ==========================================================
# PAGINATION
# ==========================================================

def get_last_page(soup):
    pages = [1]

    for a in soup.select('a[href*="page="]'):
        href = a.get("href", "")

        match = re.search(
            r"[?&]page=(\d+)",
            href,
        )

        if match:
            pages.append(int(match.group(1)))

    return max(pages)


# ==========================================================
# PRODUCT SKU
# ==========================================================

def get_product_sku(product_url):
    """
    Артикул отсутствует в новой карточке каталога.
    Поэтому открываем страницу товара и читаем
    .dtop-code__value.
    """

    if not product_url:
        return ""

    if product_url in SKU_CACHE:
        return SKU_CACHE[product_url]

    sku = ""

    for attempt in range(1, 4):
        soup = get_soup(product_url, attempts=1)

        sku_el = soup.select_one(
            ".product-data .dtop-code__value"
        )

        if not sku_el:
            sku_el = soup.select_one(
                ".dtop-code__value"
            )

        if sku_el:
            sku = clean(sku_el.get_text(" ", strip=True))

        if sku:
            break

        # Повторяем запрос, если страница временно
        # не загрузилась или артикул не обнаружен.
        if attempt < 3:
            time.sleep(attempt)

    SKU_CACHE[product_url] = sku

    return sku


# ==========================================================
# PARSE CATEGORY
# ==========================================================

def parse_category(category_url):
    all_items = []

    first_page = get_soup(category_url)

    if not first_page.select_one("body"):
        print(f"⚠️ Не загрузилась категория: {category_url}")
        return all_items

    last_page = get_last_page(first_page)

    print(
        f"📄 Страниц: {last_page} | "
        f"{category_url}"
    )

    for page in range(1, last_page + 1):

        if page == 1:
            soup = first_page
        else:
            page_url = build_page_url(
                category_url,
                page,
            )
            soup = get_soup(page_url)

        cards = soup.select("div.product-thumb")

        print(
            f"   Страница {page}/{last_page}: "
            f"{len(cards)} карточек"
        )

        for card in cards:

            # TITLE + URL
            title_el = card.select_one(
                ".product-thumb__name"
            )

            if not title_el:
                continue

            title = clean(
                title_el.get_text(" ", strip=True)
            )

            product_url = absolute_url(
                title_el.get("href", ""),
                category_url,
            )

            if not product_url:
                image_link = card.select_one(
                    ".product-thumb__image a[href]"
                )

                if image_link:
                    product_url = absolute_url(
                        image_link.get("href", ""),
                        category_url,
                    )

            if not title or not product_url:
                continue

            # PRICE
            price_el = card.select_one(
                ".product-thumb__price"
            )

            price = ""

            if price_el:
                price = clean(
                    price_el.get_text(" ", strip=True)
                )

                if not price:
                    price = clean(
                        price_el.get("data-price", "")
                    )

            # STATUS
            status_el = card.select_one(
                ".qty-indicator__text"
            )

            status = (
                clean(status_el.get_text(" ", strip=True))
                if status_el
                else ""
            )

            # SKU — со страницы товара.
            sku = get_product_sku(product_url)

            all_items.append([
                sku,
                title,
                price,
                status,
                product_url,
            ])

            # Небольшая пауза между запросами страниц
            # товаров, чтобы не создавать лишнюю нагрузку.
            time.sleep(0.05)

    return all_items


# ==========================================================
# MAIN
# ==========================================================

def run_parser():

    if is_locked():
        print("⚠️ D-Top уже запущен")
        return

    set_lock(True)

    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        save_status(
            True,
            0,
            USER,
            FILE_PATH,
        )

        if not login():
            save_status(
                False,
                0,
                USER,
                FILE_PATH,
            )
            return

        # Собираем категории.
        categories = get_categories()

        if CATEGORY_LIMIT:
            categories = categories[:CATEGORY_LIMIT]

        total_categories = len(categories)

        if total_categories == 0:
            print("❌ Категории не найдены. Excel не изменён.")

            save_status(
                False,
                0,
                USER,
                FILE_PATH,
            )
            return

        wb = Workbook()
        ws = wb.active
        ws.title = "D-Top"

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL",
        ])

        seen = set()
        total_products = 0

        print(
            f"🚀 Начинаем обработку "
            f"{total_categories} категорий"
        )

        for index, category_url in enumerate(
            categories,
            start=1,
        ):

            progress = int(
                (index - 1) / total_categories * 100
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH,
            )

            try:
                items = parse_category(category_url)

            except Exception as exc:
                print(
                    f"⚠️ Ошибка категории: "
                    f"{category_url} | "
                    f"{type(exc).__name__}"
                )
                continue

            for sku, title, price, status, url in items:

                # Не используем только title + price:
                # разные товары могут иметь одинаковые названия
                # и цены. Предпочитаем URL товара.
                key = url or (sku, title)

                if key in seen:
                    continue

                seen.add(key)

                if not title:
                    continue

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    url,
                ])

                total_products += 1

            save_status(
                True,
                int(index / total_categories * 100),
                USER,
                FILE_PATH,
            )

            print(
                f"📊 Категория {index}/{total_categories} | "
                f"Уникальных товаров: {total_products}"
            )

            time.sleep(0.1)

        # Не заменяем существующий файл, если сбор
        # фактически не дал ни одного товара.
        if total_products == 0:
            print(
                "❌ Не собрано ни одного товара. "
                "Существующий Excel сохранён без изменений."
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH,
            )
            return

        tmp_path = FILE_PATH + ".tmp"

        wb.save(tmp_path)
        os.replace(tmp_path, FILE_PATH)

        save_status(
            False,
            100,
            USER,
            FILE_PATH,
        )

        print("=" * 50)
        print("✅ D-TOP: ПАРСИНГ ЗАВЕРШЁН")
        print(f"📂 Категорий: {total_categories}")
        print(f"📦 Товаров: {total_products}")
        print(f"📄 Файл: {FILE_PATH}")
        print("=" * 50)

    except Exception as exc:
        print(
            f"❌ Критическая ошибка: "
            f"{type(exc).__name__}: {exc}"
        )

        save_status(
            False,
            0,
            USER,
            FILE_PATH,
        )

    finally:
        set_lock(False)


if __name__ == "__main__":
    run_parser()
