
import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
import sys


# =========================================================
# 👤 USER
# =========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"


print("🔥 DELLTA LIFE PARSER")


# =========================================================
# ⚙️ CONFIG
# =========================================================

BASE = "https://b2b.delltalife.com"

# Тестируем сначала 2 категории
CATEGORY_LIMIT = 2

# После успешного теста:
# CATEGORY_LIMIT = None


EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"


OUTPUT_DIR = os.path.abspath(
    "output/Dellta"
)

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Dellta_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)


# =========================================================
# 🌐 HEADERS
# =========================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),

    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,"
        "image/avif,image/webp,"
        "*/*;q=0.8"
    ),

    "Accept-Language": (
        "uk-UA,uk;q=0.9,"
        "ru;q=0.8,en;q=0.7"
    ),

    "Accept-Encoding": (
        "gzip, deflate"
    ),

    "Connection": "keep-alive",

    "Upgrade-Insecure-Requests": "1"
}


# =========================================================
# 🔗 SESSION
# =========================================================

session = requests.Session()

session.headers.update(
    HEADERS
)


# =========================================================
# STATUS
# =========================================================

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


# =========================================================
# 🧹 CLEAN
# =========================================================

def clean(text):

    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        text
    ).strip()


# =========================================================
# 🌐 HTTP REQUEST
# =========================================================

def request_page(
    url,
    method="GET",
    data=None,
    referer=None
):

    headers = dict(
        HEADERS
    )

    if referer:

        headers["Referer"] = referer

    try:

        if method == "POST":

            r = session.post(
                url,
                data=data,
                headers=headers,
                timeout=(15, 60),
                allow_redirects=True
            )

        else:

            r = session.get(
                url,
                headers=headers,
                timeout=(15, 60),
                allow_redirects=True
            )

        print(
            f"🌐 {method} {url} → HTTP {r.status_code}"
        )

        # =================================================
        # 429
        # =================================================

        if r.status_code == 429:

            print("")
            print(
                "❌ Dellta вернул HTTP 429"
            )

            print(
                "❌ Ответ: Захищена сторінка"
            )

            print(
                "🌐 Final URL:",
                r.url
            )

            print(
                "📄 Content-Type:",
                r.headers.get(
                    "Content-Type",
                    ""
                )
            )

            location = r.headers.get(
                "Location"
            )

            if location:

                print(
                    "↪ Location:",
                    location
                )

            set_cookie = r.headers.get(
                "Set-Cookie",
                ""
            )

            if set_cookie:

                print(
                    "🍪 Set-Cookie:",
                    set_cookie[:500]
                )

            print(
                "📄 Response title:"
            )

            try:

                soup = BeautifulSoup(
                    r.text,
                    "html.parser"
                )

                title = soup.title

                if title:

                    print(
                        "   ",
                        clean(
                            title.get_text()
                        )
                    )

            except Exception:
                pass

            print("")

            return None

        # =================================================
        # OTHER HTTP ERRORS
        # =================================================

        if r.status_code >= 400:

            print(
                f"❌ HTTP ERROR: {r.status_code}"
            )

            print(
                r.text[:500]
            )

            return None

        return r

    except requests.exceptions.Timeout:

        print(
            f"⏱ TIMEOUT: {url}"
        )

        return None

    except requests.exceptions.ConnectionError as e:

        print(
            f"🔌 CONNECTION ERROR: {e}"
        )

        return None

    except requests.exceptions.RequestException as e:

        print(
            f"❌ REQUEST ERROR: {e}"
        )

        return None


# =========================================================
# 🏠 INITIAL CONNECTION
# =========================================================

def open_main_page():

    print("")
    print(
        "🌐 Проверяем соединение с Dellta..."
    )

    r = request_page(
        BASE,
        method="GET"
    )

    if not r:

        print(
            "❌ Главная страница Dellta недоступна"
        )

        return False

    print(
        "✅ Главная страница получена"
    )

    print(
        "📏 Размер ответа:",
        len(r.text),
        "байт"
    )

    return True


# =========================================================
# 🔐 LOGIN
# =========================================================

def login():

    print("")
    print(
        "🔐 Авторизация Dellta..."
    )

    # -----------------------------------------------------
    # Сначала открываем главную.
    # Получаем обычную HTTP-сессию.
    # -----------------------------------------------------

    r_main = request_page(
        BASE,
        method="GET"
    )

    if not r_main:

        print(
            "❌ Не удалось открыть главную страницу"
        )

        return False

    print(
        "✅ Главная страница открыта"
    )

    # -----------------------------------------------------
    # LOGIN ENDPOINT
    # -----------------------------------------------------

    login_action = (
        BASE
        + "/themes/default/ajax/login.php"
    )

    payload = {
        "email_auth": EMAIL,
        "pass_auth": PASSWORD
    }

    headers = {
        "User-Agent": HEADERS["User-Agent"],

        "Accept": "*/*",

        "Accept-Language": (
            "uk-UA,uk;q=0.9,"
            "ru;q=0.8,en;q=0.7"
        ),

        "Content-Type": (
            "application/x-www-form-urlencoded; "
            "charset=UTF-8"
        ),

        "X-Requested-With": "XMLHttpRequest",

        "Origin": BASE,

        "Referer": BASE + "/",

        "Connection": "keep-alive"
    }

    print(
        "🔐 Отправляем login.php..."
    )

    try:

        r = session.post(
            login_action,
            data=payload,
            headers=headers,
            timeout=(15, 60),
            allow_redirects=True
        )

    except requests.exceptions.RequestException as e:

        print(
            "❌ LOGIN ERROR:",
            e
        )

        return False

    print(
        f"🔐 LOGIN HTTP: {r.status_code}"
    )

    # =====================================================
    # 429
    # =====================================================

    if r.status_code == 429:

        print("")
        print(
            "❌ LOGIN заблокирован Dellta: HTTP 429"
        )

        print(
            "❌ Сервер вернул «Захищена сторінка»"
        )

        print(
            "🌐 Final URL:",
            r.url
        )

        print(
            "📄 Response:"
        )

        print(
            r.text[:700]
        )

        print("")

        return False

    # =====================================================
    # OTHER ERRORS
    # =====================================================

    if r.status_code >= 400:

        print(
            f"❌ LOGIN HTTP ERROR: {r.status_code}"
        )

        print(
            r.text[:700]
        )

        return False

    # =====================================================
    # RESPONSE
    # =====================================================

    response_text = clean(
        r.text
    )

    print(
        "🔐 LOGIN RESPONSE:"
    )

    print(
        response_text[:500]
    )

    # =====================================================
    # COOKIE CHECK
    # =====================================================

    print("")
    print(
        "🍪 SESSION COOKIES:"
    )

    for cookie in session.cookies:

        print(
            f"   {cookie.name}"
        )

    print("")

    # =====================================================
    # LOGIN RESULT
    # =====================================================

    lower = response_text.lower()

    if (
        "error" in lower
        or "помил" in lower
        or "ошиб" in lower
    ):

        print(
            "⚠️ Сервер вернул возможную ошибку авторизации"
        )

        return False

    print(
        "✅ LOGIN REQUEST OK"
    )

    return True


# =========================================================
# 📄 GET SOUP
# =========================================================

def get_soup(url):

    r = request_page(
        url,
        method="GET"
    )

    if not r:

        return None

    return BeautifulSoup(
        r.text,
        "html.parser"
    )


# =========================================================
# 📂 CATEGORIES
# =========================================================

def get_categories():

    print("")
    print(
        "📂 Получаем категории..."
    )

    r = request_page(
        BASE,
        method="GET"
    )

    if not r:

        return []

    soup = BeautifulSoup(
        r.text,
        "html.parser"
    )

    categories = []

    container = soup.select_one(
        "div.brandsOnMain"
    )

    if not container:

        print(
            "❌ div.brandsOnMain не найден"
        )

        return categories

    for a in container.select(
        "a.COMitem"
    ):

        href = a.get(
            "href"
        )

        if not href:

            continue

        if href.startswith("/"):

            href = (
                BASE.rstrip("/")
                + href
            )

        elif not href.startswith(
            "http"
        ):

            href = (
                BASE.rstrip("/")
                + "/"
                + href.lstrip("/")
            )

        if href not in categories:

            categories.append(
                href
            )

    print(
        f"📂 Найдено категорий: {len(categories)}"
    )

    return categories


# =========================================================
# 📄 LAST PAGE
# =========================================================

def get_last_page(soup):

    if not soup:

        return 1

    pages = []

    for a in soup.select(
        ".pagination .page-link[pn]"
    ):

        pn = a.get(
            "pn"
        )

        if pn and pn.isdigit():

            pages.append(
                int(pn)
            )

    return (
        max(pages)
        if pages
        else 1
    )


# =========================================================
# 🛒 PRODUCT CARD
# =========================================================

def parse_product_card(card):

    title = ""
    sku = ""
    status = ""
    price = ""
    url = ""

    # =====================================================
    # SKU
    # =====================================================

    sku_el = card.select_one(
        ".td_2 .gray"
    )

    if sku_el:

        sku = clean(
            sku_el.get_text(
                " ",
                strip=True
            )
        )

    # =====================================================
    # TITLE + URL
    # =====================================================

    title_el = card.select_one(
        "td.td_2 a[href]"
    )

    if title_el:

        title = clean(
            title_el.get_text(
                " ",
                strip=True
            )
        )

        href = title_el.get(
            "href",
            ""
        )

        if href:

            if href.startswith("/"):

                url = (
                    BASE.rstrip("/")
                    + href
                )

            elif href.startswith(
                "http"
            ):

                url = href

            else:

                url = (
                    BASE.rstrip("/")
                    + "/"
                    + href.lstrip("/")
                )

    # =====================================================
    # FALLBACK URL
    # =====================================================

    if not url:

        title_el = card.select_one(
            'td.td_2 a[href*="/invertoryi-"]'
        )

        if title_el:

            href = title_el.get(
                "href",
                ""
            )

            if href.startswith("/"):

                url = (
                    BASE.rstrip("/")
                    + href
                )

            elif href.startswith(
                "http"
            ):

                url = href

    # =====================================================
    # STATUS
    # =====================================================

    status_el = card.select_one(
        "td.td_2 .are-available"
    )

    if status_el:

        status = clean(
            status_el.get_text(
                " ",
                strip=True
            )
        )

    else:

        status_el = card.select_one(
            "td.td_2 div[class^='are-']"
        )

        if status_el:

            status = clean(
                status_el.get_text(
                    " ",
                    strip=True
                )
            )

    # =====================================================
    # DEALER PRICE
    # =====================================================

    price_el = card.select_one(
        "td.td_3 "
        ".price-table "
        "tr.line-1 "
        "span.active"
    )

    if price_el:

        price = clean(
            price_el.get_text(
                " ",
                strip=True
            )
        )

    # =====================================================
    # FALLBACK PRICE
    # =====================================================

    if not price:

        price_el = card.select_one(
            "td.td_3 "
            "tr.line-1 "
            "span.active"
        )

        if price_el:

            price = clean(
                price_el.get_text(
                    " ",
                    strip=True
                )
            )

    return [
        sku,
        title,
        price,
        status,
        url
    ]


# =========================================================
# 📦 PARSE CATEGORY
# =========================================================

def parse_category(cat_url):

    all_items = []

    print("")
    print(
        "📂 CATEGORY:",
        cat_url
    )

    first_page = get_soup(
        cat_url
    )

    if not first_page:

        print(
            "❌ Не удалось получить категорию"
        )

        return []

    last_page = get_last_page(
        first_page
    )

    print(
        f"📄 Страниц: {last_page}"
    )

    for page in range(
        1,
        last_page + 1
    ):

        print(
            f"   📄 Страница {page}/{last_page}"
        )

        if page == 1:

            soup = first_page

        else:

            page_url = (
                cat_url.rstrip("/")
                + f"/page={page}/"
            )

            soup = get_soup(
                page_url
            )

            if not soup:

                print(
                    "   ❌ Страница не получена"
                )

                continue

        cards = soup.select(
            "tr.itemPosition"
        )

        print(
            f"   🛒 Товаров на странице: {len(cards)}"
        )

        for card in cards:

            item = parse_product_card(
                card
            )

            sku, title, price, status, url = item

            if not title:

                continue

            all_items.append(
                item
            )

    return all_items


# =========================================================
# 🚀 MAIN
# =========================================================

def run_parser():

    save_status(
        True,
        0,
        USER,
        FILE_PATH
    )

    print("")
    print(
        "🚀 Запуск парсера Dellta"
    )

    # =====================================================
    # LOGIN
    # =====================================================

    if not login():

        save_status(
            False,
            0,
            USER,
            ""
        )

        print("")
        print(
            "❌ LOGIN FAILED"
        )

        return

    # =====================================================
    # EXCEL
    # =====================================================

    wb = Workbook()

    ws = wb.active

    ws.title = "Dellta"

    ws.append([
        "SKU",
        "TITLE",
        "PRICE",
        "STATUS",
        "URL"
    ])

    # =====================================================
    # CATEGORIES
    # =====================================================

    cats = get_categories()

    if not cats:

        print("")
        print(
            "❌ CATEGORIES NOT FOUND"
        )

        save_status(
            False,
            0,
            USER,
            ""
        )

        return

    if CATEGORY_LIMIT:

        cats = cats[
            :CATEGORY_LIMIT
        ]

        print(
            f"⚙️ Тестовый лимит категорий: {CATEGORY_LIMIT}"
        )

    total = len(cats)

    print("")
    print(
        f"📂 Будет обработано категорий: {total}"
    )

    # =====================================================
    # PARSE
    # =====================================================

    for i, cat in enumerate(
        cats,
        1
    ):

        progress = int(
            i / total * 100
        )

        save_status(
            True,
            progress,
            USER,
            FILE_PATH
        )

        print("")
        print(
            f"🔥 CATEGORY {i}/{total}"
        )

        items = parse_category(
            cat
        )

        print(
            f"✅ Получено товаров: {len(items)}"
        )

        for item in items:

            sku, title, price, status, url = item

            ws.append([
                sku,
                title,
                price,
                status,
                url
            ])

        time.sleep(
            0.3
        )

    # =====================================================
    # SAVE
    # =====================================================

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    tmp = FILE_PATH + ".tmp"

    wb.save(
        tmp
    )

    os.replace(
        tmp,
        FILE_PATH
    )

    # =====================================================
    # DONE
    # =====================================================

    save_status(
        False,
        100,
        USER,
        FILE_PATH
    )

    print("")
    print(
        "===================================="
    )

    print(
        "✅ DELLTA PARSER DONE"
    )

    print(
        "📄 FILE:",
        FILE_PATH
    )

    print(
        "===================================="
    )


# =========================================================
# START
# =========================================================

if __name__ == "__main__":

    run_parser()
