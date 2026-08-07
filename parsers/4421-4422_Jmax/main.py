import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Харьковская 4421-4422 Jmax")

BASE = "https://www.jmaxtvshop.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

# =========================
# HEADERS
# =========================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,image/avif,image/webp,"
        "image/apng,*/*;q=0.8"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
    "Connection": "keep-alive"
}

session = requests.Session()

session.headers.update(HEADERS)


# =========================
# LOCK
# =========================

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:

        age = time.time() - os.path.getmtime(LOCK_FILE)

        if age > 3600:

            os.remove(LOCK_FILE)

            return False

        return True

    except:

        return False


def set_lock(state):

    if state:

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        with open(
            LOCK_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(str(time.time()))

    else:

        if os.path.exists(LOCK_FILE):

            os.remove(LOCK_FILE)


# =========================
# STATUS
# =========================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path=""
):

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        "file_path": file_path
    }

    tmp = STATUS_PATH + ".tmp"

    with open(
        tmp,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )

    os.replace(
        tmp,
        STATUS_PATH
    )


# =========================
# LOGIN
# =========================

def login():

    print("")
    print("🔐 LOGIN...", flush=True)

    login_url = (
        BASE +
        "/index.php?route=account/login"
    )

    print(
        f"🔐 URL: {login_url}",
        flush=True
    )

    try:

        # ---------------------------------
        # Получаем страницу логина
        # ---------------------------------

        r = session.get(
            login_url,
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔐 LOGIN GET STATUS: {r.status_code}",
            flush=True
        )

        print(
            f"🔐 LOGIN GET URL: {r.url}",
            flush=True
        )

        print(
            f"🔐 LOGIN HTML: {len(r.text)}",
            flush=True
        )

        # ---------------------------------
        # 429
        # ---------------------------------

        if r.status_code == 429:

            print(
                "❌ Jmax вернул HTTP 429.",
                flush=True
            )

            print(
                "⚠️ Railway IP сейчас ограничен сайтом.",
                flush=True
            )

            return False

        if r.status_code != 200:

            print(
                f"❌ LOGIN GET ERROR: HTTP {r.status_code}",
                flush=True
            )

            return False

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        # ---------------------------------
        # Ищем форму
        # ---------------------------------

        form = soup.select_one(
            'form[action*="route=account/login"]'
        )

        if not form:

            form = soup.select_one("form")

        if not form:

            print(
                "❌ Форма логина не найдена.",
                flush=True
            )

            return False

        print(
            "✅ Форма логина найдена.",
            flush=True
        )

        # ---------------------------------
        # Собираем hidden-поля
        # ---------------------------------

        payload = {}

        for inp in form.select("input"):

            name = inp.get("name")

            if not name:
                continue

            payload[name] = inp.get(
                "value",
                ""
            )

        # ---------------------------------
        # Логин / пароль
        # ---------------------------------

        payload["email"] = EMAIL
        payload["password"] = PASSWORD

        print(
            "🔐 Отправляем POST LOGIN...",
            flush=True
        )

        # ---------------------------------
        # POST
        # ---------------------------------

        response = session.post(
            login_url,
            data=payload,
            headers={
                "Referer": login_url,
                "User-Agent": HEADERS["User-Agent"]
            },
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔐 LOGIN POST STATUS: "
            f"{response.status_code}",
            flush=True
        )

        print(
            f"🔐 LOGIN POST URL: "
            f"{response.url}",
            flush=True
        )

        # ---------------------------------
        # 429
        # ---------------------------------

        if response.status_code == 429:

            print(
                "❌ POST LOGIN получил HTTP 429.",
                flush=True
            )

            return False

        # ---------------------------------
        # Проверяем ответ
        # ---------------------------------

        html = response.text.lower()

        if (
            "account/logout" in html
            or "route=account/logout" in html
            or "logout" in html
            or "выйти" in html
            or "вихід" in html
        ):

            print(
                "✅ LOGIN OK",
                flush=True
            )

            return True

        # ---------------------------------
        # Проверяем страницу аккаунта
        # ---------------------------------

        if (
            "account/account"
            in response.url
        ):

            print(
                "✅ LOGIN OK",
                flush=True
            )

            return True

        # ---------------------------------
        # Дополнительная проверка
        # ---------------------------------

        print(
            "🔐 Проверяем авторизацию через главную...",
            flush=True
        )

        account = session.get(
            BASE,
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔐 MAIN STATUS: "
            f"{account.status_code}",
            flush=True
        )

        if account.status_code == 429:

            print(
                "❌ Главная также вернула HTTP 429.",
                flush=True
            )

            return False

        account_html = account.text.lower()

        if (
            "account/logout" in account_html
            or "route=account/logout" in account_html
            or "logout" in account_html
            or "выйти" in account_html
            or "вихід" in account_html
        ):

            print(
                "✅ LOGIN OK",
                flush=True
            )

            return True

        print(
            "⚠️ LOGIN CHECK — не удалось подтвердить вход.",
            flush=True
        )

        return False

    except Exception as e:

        print(
            f"❌ LOGIN ERROR: {e}",
            flush=True
        )

        return False


# =========================
# HTTP
# =========================

def get_soup(url):

    for attempt in range(1, 4):

        try:

            r = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            if r.status_code == 200:

                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

            if r.status_code == 429:

                print(
                    f"⚠️ HTTP 429: {url}",
                    flush=True
                )

                return BeautifulSoup(
                    "",
                    "html.parser"
                )

            print(
                f"⚠️ HTTP {r.status_code}: {url}",
                flush=True
            )

        except Exception as e:

            print(
                f"⚠️ GET ERROR: {e}",
                flush=True
            )

        time.sleep(attempt)

    return BeautifulSoup(
        "",
        "html.parser"
    )


# =========================
# CLEAN
# =========================

def clean(t):

    return (
        re.sub(
            r"\s+",
            " ",
            t
        ).strip()
        if t
        else ""
    )


# =========================
# CATEGORIES
# =========================

def get_categories():

    print("")
    print("=" * 70)
    print("🌳 ПОИСК ВСЕХ КАТЕГОРИЙ JMAX")
    print("=" * 70)

    soup = get_soup(BASE)

    if not soup or not soup.find_all(True):

        print(
            "❌ Главная страница не загрузилась",
            flush=True
        )

        return []

    categories = []
    seen = set()

    # ---------------------------------
    # Нормализация URL
    # ---------------------------------

    def normalize_url(url):

        if not url:

            return ""

        url = url.strip()

        if not url:

            return ""

        if url.startswith("#"):

            return ""

        if url.startswith(
            "javascript:"
        ):

            return ""

        if not url.startswith(
            "http"
        ):

            url = (
                BASE +
                "/" +
                url.lstrip("/")
            )

        if (
            "route=product/category"
            not in url
        ):

            return ""

        # убираем page
        url = re.sub(
            r"[&?]page=\d+",
            "",
            url
        )

        return url

    # ---------------------------------
    # Добавление категории
    # ---------------------------------

    def add_category(url):

        url = normalize_url(url)

        if not url:

            return

        if url in seen:

            return

        seen.add(url)

        categories.append(url)

        print(
            f"📂 Категория #{len(categories)}: "
            f"{url}",
            flush=True
        )

    # ---------------------------------
    # Сканируем ВСЕ ссылки
    # ---------------------------------

    for a in soup.find_all(
        "a",
        href=True
    ):

        add_category(
            a.get("href")
        )

    # ---------------------------------
    # data-href
    # ---------------------------------

    for tag in soup.find_all(
        attrs={"data-href": True}
    ):

        add_category(
            tag.get("data-href")
        )

    print("")
    print(
        f"🌳 Найдено категорий: "
        f"{len(categories)}",
        flush=True
    )

    # ---------------------------------
    # Рекурсивно ищем скрытые категории
    # ---------------------------------

    index = 0

    while index < len(categories):

        cat_url = categories[index]

        index += 1

        print(
            f"🌳 [{index}/{len(categories)}] "
            f"Сканируем категорию:",
            cat_url,
            flush=True
        )

        soup = get_soup(cat_url)

        if not soup or not soup.find_all(True):

            continue

        before = len(categories)

        for a in soup.find_all(
            "a",
            href=True
        ):

            add_category(
                a.get("href")
            )

        for tag in soup.find_all(
            attrs={"data-href": True}
        ):

            add_category(
                tag.get("data-href")
            )

        new_count = (
            len(categories) - before
        )

        if new_count:

            print(
                f"   ➕ Новых категорий: "
                f"{new_count}",
                flush=True
            )

    print("")
    print("=" * 70)
    print(
        f"🌳 ВСЕГО КАТЕГОРИЙ: "
        f"{len(categories)}",
        flush=True
    )
    print("=" * 70)

    return categories


# =========================
# LAST PAGE
# =========================

def get_last_page(soup):

    pages = [1]

    for a in soup.select(
        "ul.pagination a"
    ):

        href = a.get(
            "href",
            ""
        )

        m = re.search(
            r"[?&]page=(\d+)",
            href
        )

        if m:

            pages.append(
                int(m.group(1))
            )

    return max(pages)


# =========================
# PRODUCT
# =========================

def parse_product_soup(
    soup,
    url
):

    # TITLE

    title = ""

    h1 = soup.select_one("h1")

    if h1:

        title = clean(
            h1.get_text()
        )

    # SKU

    sku = ""

    sku_tag = soup.select_one(
        ".product-data__item.model"
    )

    if sku_tag:

        sku = clean(
            sku_tag.get_text()
            .replace(
                "Код товара:",
                ""
            )
        )

    # PRICE

    price = ""

    p = soup.select_one(
        ".product-page__price"
    )

    if p:

        price = clean(
            p.get_text()
        )

    # STATUS

    status = ""

    btn = soup.select_one(
        "#button-cart span"
    )

    if btn:

        status = clean(
            btn.get_text()
        )

    else:

        status = "Нет кнопки"

    return [
        sku,
        title,
        price,
        status,
        url
    ]


def parse_product(url):

    soup = get_soup(url)

    if not soup.select_one("h1"):

        return [
            "",
            "",
            "",
            "",
            url
        ]

    return parse_product_soup(
        soup,
        url
    )


# =========================
# PARSE CATEGORY
# =========================

def parse_category(cat_url):

    all_items = []

    # ---------------------------------
    # Первая страница
    # ---------------------------------

    first_page = get_soup(
        cat_url
    )

    if not first_page or not first_page.find_all(True):

        return []

    last_page = get_last_page(
        first_page
    )

    print(
        f"📄 Страниц в категории: "
        f"{last_page}",
        flush=True
    )

    # ---------------------------------
    # Все страницы
    # ---------------------------------

    seen_products = set()

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:

            url = cat_url

        else:

            url = (
                f"{cat_url}&page={page}"
            )

        print(
            f"📄 Страница "
            f"{page}/{last_page}",
            flush=True
        )

        soup = get_soup(
            url
        )

        if not soup or not soup.find_all(True):

            continue

        products = []

        # ---------------------------------
        # Ищем ссылки товаров
        # ---------------------------------

        for a in soup.find_all(
            "a",
            href=True
        ):

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:

                continue

            if not href.startswith(
                "http"
            ):

                href = (
                    BASE +
                    "/" +
                    href.lstrip("/")
                )

            if (
                "route=product/product"
                not in href
            ):

                continue

            href = href.split("#")[0]

            if href in products:

                continue

            products.append(
                href
            )

        # ---------------------------------
        # Нет товаров
        # ---------------------------------

        if not products:

            print(
                "   ⛔ Товаров нет",
                flush=True
            )

            continue

        print(
            f"   📦 Найдено товаров: "
            f"{len(products)}",
            flush=True
        )

        # ---------------------------------
        # Парсим товары
        # ---------------------------------

        for href in products:

            if href in seen_products:

                continue

            seen_products.add(
                href
            )

            item = parse_product(
                href
            )

            if item and item[1]:

                all_items.append(
                    item
                )

            time.sleep(0.05)

    print(
        f"✅ Категория закончена: "
        f"{len(all_items)} товаров",
        flush=True
    )

    return all_items


# =========================
# MAIN
# =========================

def run_parser():

    if is_locked():

        print(
            "⚠️ Парсер уже запущен",
            flush=True
        )

        return

    set_lock(True)

    try:

        save_status(
            True,
            0,
            USER,
            FILE_PATH
        )

        # ---------------------------------
        # LOGIN
        # ---------------------------------

        if not login():

            print(
                "❌ Не удалось авторизоваться",
                flush=True
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # ---------------------------------
        # EXCEL
        # ---------------------------------

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        # ---------------------------------
        # CATEGORIES
        # ---------------------------------

        cats = get_categories()

        print(
            f"🌳 Получено категорий: "
            f"{len(cats)}",
            flush=True
        )

        if CATEGORY_LIMIT:

            cats = cats[
                :CATEGORY_LIMIT
            ]

        total_categories = len(cats)

        if total_categories == 0:

            print(
                "❌ Категории не найдены",
                flush=True
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # ---------------------------------
        # PRODUCTS
        # ---------------------------------

        seen = set()

        for i, cat in enumerate(
            cats,
            1
        ):

            progress = int(
                (i - 1)
                /
                total_categories
                *
                100
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )

            print("")
            print("=" * 70)

            print(
                f"📂 КАТЕГОРИЯ "
                f"{i}/{total_categories}",
                flush=True
            )

            print(
                cat,
                flush=True
            )

            print("=" * 70)

            items = parse_category(
                cat
            )

            for item in items:

                sku = item[0]
                title = item[1]
                price = item[2]
                status = item[3]
                url = item[4]

                if not title:

                    continue

                # ---------------------------------
                # ДЕДУПЛИКАЦИЯ
                # ---------------------------------

                key = (
                    sku.strip()
                    if sku
                    else url.strip()
                )

                if not key:

                    key = (
                        title,
                        price
                    )

                if key in seen:

                    continue

                seen.add(key)

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    url
                ])

            time.sleep(0.2)

        # ---------------------------------
        # SAVE
        # ---------------------------------

        print("")
        print("=" * 70)
        print("💾 СОХРАНЕНИЕ EXCEL")
        print("=" * 70)

        tmp = FILE_PATH + ".tmp"

        wb.save(tmp)

        os.replace(
            tmp,
            FILE_PATH
        )

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        print("")
        print("=" * 70)
        print(
            "✅ ГОТОВО. "
            "Харьковская 4421-4422 Jmax",
            flush=True
        )

        print(
            f"📊 Всего товаров: "
            f"{len(seen)}",
            flush=True
        )

        print(
            f"📂 Всего категорий: "
            f"{total_categories}",
            flush=True
        )

        print("=" * 70)

    finally:

        set_lock(False)


# =========================
# START
# =========================

if __name__ == "__main__":

    run_parser()
