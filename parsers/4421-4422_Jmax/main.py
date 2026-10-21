
# -*- coding: utf-8 -*-

import os
import sys
import re
import json
import time
import logging
import requests

from datetime import datetime
from urllib.parse import urljoin, urlparse, urlunparse, parse_qs, urlencode
from bs4 import BeautifulSoup
from openpyxl import Workbook


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Харьковская 4421-4422 Jmax")

BASE = "https://www.jmaxtvshop.com.ua"

CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

# Задаются в окружении, чтобы не хранить пароль в коде.
EMAIL = os.getenv("JMAX_EMAIL", "")
PASSWORD = os.getenv("JMAX_PASSWORD", "")

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")

FILE_PATH = os.path.join(
    OUTPUT_DIR, "4421-4422_Jmax_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR, "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR, "lock.txt"
)

MAX_PAGES = 150
MAX_RETRIES = 3
REQUEST_DELAY = 0.25

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:158.0) "
        "Gecko/20100101 Firefox/158.0"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,*/*;q=0.8"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}

session = requests.Session()
session.headers.update(HEADERS)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(message)s",
    datefmt="%H:%M:%S",
)


# ============================================================
# LOCK — совместим со старой интеграцией
# ============================================================

def is_locked():
    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        if age > 3600:
            os.remove(LOCK_FILE)
            return False

        return True

    except OSError:
        return True


def set_lock(state):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if state:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))
    elif os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# ============================================================
# STATUS — формат старого status.json
# ============================================================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path=""
):
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


# ============================================================
# HTTP
# ============================================================

def clean(text):
    return re.sub(r"\s+", " ", text or "").strip()


def normalize_url(url):
    if not url:
        return ""

    url = url.strip()

    if url.startswith(("#", "javascript:", "mailto:", "tel:")):
        return ""

    absolute = urljoin(BASE + "/", url)
    parsed = urlparse(absolute)

    if parsed.netloc.lower() != urlparse(BASE).netloc.lower():
        return ""

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        parsed.query,
        "",
    ))


def canonical_url(url):
    url = normalize_url(url)
    if not url:
        return ""

    parsed = urlparse(url)

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path.rstrip("/") or "/",
        "",
        "",
        "",
    ))


def get_soup(url):
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(
                url,
                timeout=30,
                allow_redirects=True,
            )

            if response.status_code == 429:
                wait = min(60 * attempt, 300)
                print(f"⚠️ Jmax 429. Ждём {wait} сек.")
                if attempt < MAX_RETRIES:
                    time.sleep(wait)
                    continue
                break

            response.raise_for_status()

            if not response.encoding or response.encoding.lower() == "iso-8859-1":
                response.encoding = response.apparent_encoding or "utf-8"

            time.sleep(REQUEST_DELAY)
            return BeautifulSoup(response.text, "html.parser")

        except requests.RequestException as exc:
            logging.warning("Ошибка GET %s: %s", url, exc)

            if attempt < MAX_RETRIES:
                time.sleep(attempt * 2)

    return BeautifulSoup("", "html.parser")


# ============================================================
# LOGIN
# ============================================================

def login():
    print("🔐 LOGIN...")

    # Если учётные данные не заданы, пробуем публичный доступ.
    if not EMAIL or not PASSWORD:
        print("ℹ️ JMAX_EMAIL/JMAX_PASSWORD не заданы; пробуем без входа.")
        return True

    login_url = BASE + "/index.php?route=account/login"

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(login_url, timeout=30)
            response.raise_for_status()

            soup = BeautifulSoup(response.text, "html.parser")

            form = soup.select_one("form.jm-login__form")

            if not form:
                for candidate in soup.select("form"):
                    if (
                        candidate.select_one('input[name="email"]')
                        and candidate.select_one('input[name="password"]')
                    ):
                        form = candidate
                        break

            if not form:
                print("❌ LOGIN FORM NOT FOUND")
                return False

            payload = {}

            for inp in form.select("input[name]"):
                input_type = (inp.get("type") or "").lower()

                if input_type in ("submit", "button", "file"):
                    continue

                payload[inp["name"]] = inp.get("value", "")

            payload["email"] = EMAIL
            payload["password"] = PASSWORD

            action = normalize_url(form.get("action") or login_url)

            result = session.post(
                action,
                data=payload,
                headers={"Referer": login_url},
                timeout=30,
                allow_redirects=True,
            )
            result.raise_for_status()

            account = session.get(BASE, timeout=30)
            account.raise_for_status()

            html = account.text.lower()

            if any(marker in html for marker in (
                "route=account/logout",
                "account/logout",
                "выйти",
                "вихід",
            )):
                print("✅ LOGIN OK")
                return True

            print("⚠️ Вход не подтверждён. Пробуем продолжить.")
            return True

        except requests.RequestException as exc:
            print(f"❌ LOGIN ERROR: {exc}")

            if attempt < MAX_RETRIES:
                time.sleep(attempt * 3)

    return False


# ============================================================
# CATEGORY DISCOVERY
# ============================================================

def get_categories():
    soup = get_soup(BASE)

    if not soup.find_all(True):
        print("❌ Главная страница не загрузилась")
        return []

    found = {}

    # Ищем ссылки именно в меню, а не по всей странице:
    # иначе ссылки на товары попадут в список категорий.
    menu_selectors = [
        "nav a[href]",
        "nav [data-href]",
        ".menu a[href]",
        ".menu [data-href]",
        ".jm-menu a[href]",
        ".jm-menu [data-href]",
        ".dropdown-menu a[href]",
        ".dropdown-menu [data-href]",
        ".header a[href]",
        ".header [data-href]",
        "header a[href]",
        "header [data-href]",
    ]

    elements = []

    for selector in menu_selectors:
        elements.extend(soup.select(selector))

    # Резервный вариант для меню, построенного на других классах.
    if not elements:
        elements = soup.select(
            "[class*='menu'] a[href], [class*='menu'] [data-href]"
        )

    blocked = (
        "account", "checkout", "route=checkout", "route=account",
        "route=product/product", "cart", "wishlist", "compare",
        "search", "login", "register", "contact", "logout",
        "manufacturer", "information",
    )

    for element in elements:
        href = element.get("href") or element.get("data-href")
        url = normalize_url(href)

        if not url:
            continue

        if any(term in url.lower() for term in blocked):
            continue

        parsed = urlparse(url)

        if parsed.path in ("", "/"):
            continue

        # Не включаем карточки товаров.
        if (
            element.select_one(".product-thumb__price")
            or element.find_parent(class_="product-thumb")
            or "product-thumb__name" in element.get("class", [])
        ):
            continue

        # Удаляем page из URL, чтобы категория не дублировалась.
        query = parse_qs(parsed.query, keep_blank_values=True)
        query.pop("page", None)

        url = urlunparse((
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            urlencode(query, doseq=True),
            "",
        ))

        label = clean(element.get_text(" ", strip=True))

        if not label:
            label = clean(element.get("title", ""))

        if not label:
            continue

        # Не добавляем служебные ссылки.
        if label.lower() in (
            "главная", "главная страница", "home",
            "войти", "регистрация", "корзина",
        ):
            continue

        found.setdefault(url, label)

    categories = list(found.items())

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

    print(f"📂 Categories: {len(categories)}")

    for url, label in categories[:10]:
        print(f"   • {label}: {url}")

    return [url for url, _ in categories]


# ============================================================
# PRODUCT LINKS
# ============================================================

def get_product_links(soup):
    products = {}

    if not soup:
        return products

    # Новый формат Jmax: SEO URL и product-thumb__name.
    for anchor in soup.select("a.product-thumb__name[href]"):
        url = canonical_url(anchor.get("href"))
        title = clean(anchor.get_text(" ", strip=True))

        if url and title:
            products[url] = title

    # Резервный селектор по карточкам каталога.
    if not products:
        for card in soup.select(".product-thumb, .uni-item"):
            anchor = card.select_one(
                "a.product-thumb__name[href], a[href][title]"
            )

            if not anchor:
                continue

            url = canonical_url(anchor.get("href"))
            title = clean(
                anchor.get_text(" ", strip=True)
                or anchor.get("title", "")
            )

            if url and title:
                products[url] = title

    return products


# ============================================================
# PAGINATION
# ============================================================

def get_last_page(soup):
    if not soup:
        return 1

    text = clean(soup.get_text(" ", strip=True))

    # Новая кнопка: "Страница 2 из 5".
    match = re.search(
        r"Страница\s+\d+\s+из\s+(\d+)",
        text,
        re.I,
    )

    if match:
        return min(int(match.group(1)), MAX_PAGES)

    pages = [1]

    for anchor in soup.select(
        "ul.pagination a[href], .pagination a[href]"
    ):
        query = parse_qs(urlparse(anchor["href"]).query)

        for value in query.get("page", []):
            if value.isdigit():
                pages.append(int(value))

    return min(max(pages), MAX_PAGES)


def page_url(category_url, page):
    parsed = urlparse(category_url)
    query = parse_qs(parsed.query, keep_blank_values=True)

    if page == 1:
        query.pop("page", None)
    else:
        query["page"] = [str(page)]

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        urlencode(query, doseq=True),
        "",
    ))


def parse_category(cat_url):
    all_products = {}

    first_page = get_soup(page_url(cat_url, 1))

    if not first_page.find_all(True):
        return all_products

    first_products = get_product_links(first_page)
    all_products.update(first_products)

    last_page = get_last_page(first_page)

    # Если число страниц не отображается в HTML,
    # перебираем страницы до первой пустой.
    known_page_count = last_page > 1

    if known_page_count:
        pages = range(2, last_page + 1)
    else:
        pages = range(2, MAX_PAGES + 1)

    for page in pages:
        soup = get_soup(page_url(cat_url, page))

        if not soup.find_all(True):
            break

        products = get_product_links(soup)

        if not products:
            break

        before = len(all_products)
        all_products.update(products)

        if not known_page_count and len(all_products) == before:
            break

    return all_products


# ============================================================
# SITEMAP — SEO URL
# ============================================================

def get_sitemap_urls():
    candidates = [
        BASE + "/sitemap.xml",
        BASE + "/sitemap_index.xml",
        BASE + "/index.php?route=feed/google_sitemap",
        BASE + "/index.php?route=extension/feed/google_sitemap",
    ]

    result = set()
    visited_sitemaps = set()

    def read_sitemap(sitemap_url, depth=0):
        if depth > 2 or sitemap_url in visited_sitemaps:
            return

        visited_sitemaps.add(sitemap_url)

        soup = get_soup(sitemap_url)

        if not soup.find_all(True):
            return

        # Поддержка sitemap index и обычного sitemap.
        for loc in soup.find_all("loc"):
            url = normalize_url(loc.get_text(strip=True))

            if not url:
                continue

            path = urlparse(url).path.lower()

            if path.endswith(".xml") or "sitemap" in path:
                read_sitemap(url, depth + 1)
            else:
                result.add(canonical_url(url))

    for candidate in candidates:
        read_sitemap(candidate)

    print(f"🗺 Sitemap URL: {len(result)}")

    return sorted(result)


# ============================================================
# PRODUCT DETAILS
# ============================================================

def parse_product(url, listing_title=""):
    soup = get_soup(url)

    if not soup.find_all(True):
        return None

    # Проверяем, что это именно товарная карточка.
    price_tag = soup.select_one(
        ".product-page__price, .jm-pp__price"
    )

    product_id = soup.select_one(
        'input[name="product_id"]'
    )

    if not price_tag and not product_id:
        return None

    h1 = soup.select_one("h1.jm-pp__title") or soup.select_one("h1")

    title = clean(h1.get_text(" ", strip=True)) if h1 else listing_title

    if not title:
        return None

    # Артикул — именно код товара, не data-pid/product_id.
    sku = ""

    model = soup.select_one(".product-data__item.model")

    if model:
        model_text = clean(model.get_text(" ", strip=True))

        match = re.search(
            r"(?:Код\s*товара|Артикул|SKU|Модель|Model|Код)"
            r"\s*[:№#-]?\s*([A-ZА-ЯЁ0-9._/-]+)",
            model_text,
            re.I,
        )

        if match:
            sku = match.group(1).strip()

    if not sku:
        page_text = clean(soup.get_text(" ", strip=True))

        match = re.search(
            r"(?:Код\s*товара|Артикул|SKU|Модель|Model)"
            r"\s*[:№#-]?\s*([A-ZА-ЯЁ0-9._/-]+)",
            page_text,
            re.I,
        )

        if match:
            sku = match.group(1).strip()

    # Цена сохраняется в отображаемом виде, например $4.00.
    price = clean(price_tag.get_text(" ", strip=True)) if price_tag else ""

    if not price and price_tag:
        price = (
            price_tag.get("content")
            or price_tag.get("data-price")
            or ""
        )

    # Новый селектор наличия.
    stock_tag = (
        soup.select_one(".jm-buy__stock")
        or soup.select_one(".product-data__item.stock")
    )

    status = clean(stock_tag.get_text(" ", strip=True)) if stock_tag else ""

    if not status:
        cart_button = soup.select_one("#button-cart span")

        if cart_button:
            status = clean(cart_button.get_text(" ", strip=True))

    return [
        sku,
        title,
        price,
        status,
        canonical_url(url),
    ]


# ============================================================
# MAIN
# ============================================================

def run_parser():
    if is_locked():
        print("⚠️ Парсер уже запущен.")
        return

    set_lock(True)

    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        save_status(True, 0, USER, FILE_PATH)

        if not login():
            print("❌ Не удалось авторизоваться")
            save_status(False, 0, USER, FILE_PATH)
            return

        wb = Workbook()
        ws = wb.active

        ws.title = "Jmax"
        ws.append(["SKU", "TITLE", "PRICE", "STATUS", "URL"])
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = "A1:E1"

        for column, width in {
            "A": 16,
            "B": 65,
            "C": 16,
            "D": 24,
            "E": 90,
        }.items():
            ws.column_dimensions[column].width = width

        product_map = {}

        # 1. Сбор товаров по категориям.
        categories = get_categories()

        if not categories:
            print(
                "❌ Категории не найдены. "
                "Проверь селекторы меню Jmax."
            )
            save_status(False, 0, USER, FILE_PATH)
            return

        for index, category in enumerate(categories, 1):
            print(
                f"📂 Категория {index}/{len(categories)}: {category}"
            )

            try:
                links = parse_category(category)
                product_map.update(links)

                print(f"   📦 Товаров найдено: {len(links)}")

            except Exception as exc:
                logging.exception(
                    "Ошибка категории %s: %s",
                    category,
                    exc,
                )

        print(f"📦 Товаров из категорий: {len(product_map)}")

        # 2. Sitemap: SEO URL вместо старого route=product/product.
        for sitemap_url in get_sitemap_urls():
            product_map.setdefault(sitemap_url, "")

        product_urls = list(product_map.items())
        total = len(product_urls)

        print(f"📦 Всего URL для проверки: {total}")

        if not total:
            print("❌ Не найдено ни одного товара.")
            save_status(False, 0, USER, FILE_PATH)
            return

        # 3. Парсинг карточек.
        errors = 0
        written = 0

        for index, (url, listing_title) in enumerate(product_urls, 1):
            progress = int(index / total * 100)

            save_status(
                True,
                progress,
                USER,
                FILE_PATH,
            )

            try:
                item = parse_product(url, listing_title)

                # Страницы категорий и служебные URL пропускаются.
                if not item:
                    continue

                sku, title, price, status, final_url = item

                if not title:
                    continue

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    final_url,
                ])

                written += 1

                # Периодическое сохранение, чтобы не терять прогресс.
                if written % 50 == 0:
                    tmp = FILE_PATH + ".tmp.xlsx"
                    wb.save(tmp)
                    os.replace(tmp, FILE_PATH)

                if written % 100 == 0:
                    print(
                        f"   📊 Обработано: {index}/{total}; "
                        f"товаров: {written}"
                    )

            except Exception as exc:
                errors += 1
                logging.warning("Ошибка товара %s: %s", url, exc)

        # 4. Финальное сохранение.
        tmp = FILE_PATH + ".tmp.xlsx"
        wb.save(tmp)
        os.replace(tmp, FILE_PATH)

        save_status(False, 100, USER, FILE_PATH)

        print(f"✅ Готово. Сохранено товаров: {written}")
        print(f"⚠️ Ошибок карточек: {errors}")
        print(f"📄 Excel: {FILE_PATH}")

    except Exception as exc:
        logging.exception("Критическая ошибка парсера: %s", exc)
        save_status(False, 0, USER, FILE_PATH)

    finally:
        set_lock(False)


if __name__ == "__main__":
    run_parser()
