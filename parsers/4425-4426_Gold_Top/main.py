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

print("🔥 Харьковская 4425-4426 Gold Top")


BASE = "https://gold-tor.com.ua"


# =========================
# ⚙️ SWITCH
# =========================

CATEGORY_LIMIT = 2

# CATEGORY_LIMIT = None


EMAIL = "Sawrun_05@icloud.com"
PASSWORD = "18022021"


OUTPUT_DIR = os.path.abspath(
    "output/4425-4426_Gold_Top"
)


FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "4425-4426_Gold_Top_LIVE.xlsx"
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
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,image/apng,*/*;"
        "q=0.8"
    ),
    "Accept-Language": (
        "uk-UA,uk;q=0.9,ru;q=0.8,en-US;q=0.7,en;q=0.6"
    ),
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}


session = requests.Session()

session.headers.update(
    HEADERS
)


# =========================
# LOCK
# =========================

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:

        age = (
            time.time()
            - os.path.getmtime(LOCK_FILE)
        )

        if age > 3600:

            os.remove(
                LOCK_FILE
            )

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

            f.write(
                str(time.time())
            )

    else:

        if os.path.exists(
            LOCK_FILE
        ):

            os.remove(
                LOCK_FILE
            )


# =========================
# COOKIES
# =========================

def print_cookies():

    # print("🍪 COOKIES:")

    try:

        for cookie in session.cookies:

            print(
                f"   {cookie.name}="
                f"{cookie.value[:20]}..."
                f" domain={cookie.domain}"
                f" path={cookie.path}"
            )

    except Exception as e:

        print(
            f"⚠️ COOKIE PRINT ERROR: {e}"
        )


def has_cookie(cookie_name):

    try:

        for cookie in session.cookies:

            if cookie.name == cookie_name:

                return True

    except Exception:

        pass

    return False


def get_cookie_values(cookie_name):

    values = []

    try:

        for cookie in session.cookies:

            if cookie.name == cookie_name:

                values.append({
                    "value": cookie.value,
                    "domain": cookie.domain,
                    "path": cookie.path
                })

    except Exception:

        pass

    return values


# =========================
# LOGIN
# =========================

def login():

    print("🔐 LOGIN...")

    login_page = (
        BASE
        + "/my-account/"
    )

    try:

        time.sleep(2)

        login_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,image/avif,"
                "image/webp,image/apng,*/*;q=0.8"
            ),
            "Accept-Language": HEADERS["Accept-Language"],
            "Accept-Encoding": "gzip, deflate",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
        }

        r = session.get(
            login_page,
            headers=login_headers,
            timeout=30,
            allow_redirects=True
        )

        if r.status_code != 200:

            print(
                f"❌ LOGIN PAGE ERROR: "
                f"{r.status_code}"
            )

            return False

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        # =========================
        # LOGIN FORM
        # =========================

        form = soup.select_one(
            "form.woocommerce-form-login"
        )

        if not form:

            form = soup.select_one(
                "form.woocommerce-form."
                "woocommerce-form-login"
            )

        if not form:

            print(
                "❌ LOGIN FORM NOT FOUND"
            )

            with open(
                os.path.join(
                    OUTPUT_DIR,
                    "login_debug.html"
                ),
                "w",
                encoding="utf-8"
            ) as f:

                f.write(
                    r.text
                )

            return False

        # =========================
        # NONCE
        # =========================

        nonce = form.select_one(
            "input[name='woocommerce-login-nonce']"
        )

        if not nonce:

            print(
                "❌ LOGIN NONCE NOT FOUND"
            )

            return False

        nonce_value = (
            nonce.get(
                "value",
                ""
            )
            .strip()
        )

        if not nonce_value:

            print(
                "❌ LOGIN NONCE EMPTY"
            )

            return False

        # =========================
        # FORM ACTION
        # =========================

        action = form.get(
            "action"
        )

        if not action:

            action = (
                login_page
                + "?action=login"
            )

        if not action.startswith(
            "http"
        ):

            action = (
                BASE.rstrip("/")
                + "/"
                + action.lstrip("/")
            )

        # =========================
        # PAYLOAD
        # =========================

        payload = {
            "username": EMAIL,
            "password": PASSWORD,
            "woocommerce-login-nonce": nonce_value,
            "_wp_http_referer": "/my-account/",
            "login": "Увійти",
            "rememberme": "forever"
        }

        time.sleep(1)

        post_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": HEADERS["Accept-Language"],
            "Accept-Encoding": "gzip, deflate",
            "Content-Type": (
                "application/x-www-form-urlencoded"
            ),
            "Referer": login_page,
            "Origin": BASE,
        }

        r = session.post(
            action,
            data=payload,
            headers=post_headers,
            allow_redirects=True,
            timeout=30
        )

        if r.status_code == 429:

            print(
                "❌ LOGIN POST: 429"
            )

            return False

        # =========================
        # ПРОВЕРКА LOGIN
        # =========================

        response_soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        response_text = (
            response_soup.get_text(
                " ",
                strip=True
            ).lower()
        )

        success_markers = [
            "вийти",
            "вихід",
            "выйти",
            "logout",
            "мій акаунт",
            "мій обліковий запис",
            "my account"
        ]

        for marker in success_markers:

            if marker in response_text:

                print(
                    f"✅ LOGIN OK "
                    f"(marker: {marker})"
                )

                return True

        # =========================
        # LOGIN FAIL
        # =========================

        print(
            "❌ LOGIN FAIL"
        )

        debug_path = os.path.join(
            OUTPUT_DIR,
            "login_after_post.html"
        )

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        with open(
            debug_path,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(
                r.text
            )

        return False

    except Exception as e:

        print(
            f"❌ LOGIN ERROR: {e}"
        )

        return False


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
# HTTP
# =========================

def get_soup(url, referer=None):

    for attempt in range(3):

        try:

            headers = {
                "User-Agent": HEADERS["User-Agent"],
                "Accept": (
                    "text/html,application/xhtml+xml,"
                    "application/xml;q=0.9,*/*;q=0.8"
                ),
                "Accept-Language": HEADERS["Accept-Language"],
                "Accept-Encoding": "identity",
                "Connection": "keep-alive",
                "Upgrade-Insecure-Requests": "1",
            }

            if referer:
                headers["Referer"] = referer

            r = session.get(
                url,
                headers=headers,
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
                    f"⚠️ 429: {url}"
                )

                time.sleep(
                    5 * (attempt + 1)
                )

                continue

            print(
                f"⚠️ HTTP {r.status_code}: {url}"
            )

        except Exception as e:

            print(
                f"⚠️ GET ERROR: {e}"
            )

            time.sleep(2)

    return BeautifulSoup(
        "",
        "html.parser"
    )



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

    print(
        "📂 ПОИСК КАТЕГОРИЙ..."
    )

    url = BASE + "/"

    try:

        category_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": HEADERS["Accept-Language"],
            "Referer": BASE + "/my-account/",
            "Connection": "keep-alive",
            "Upgrade-Insecure-Requests": "1",
            "Accept-Encoding": "identity",
        }

        time.sleep(1)

        r = session.get(
            url,
            headers=category_headers,
            timeout=30,
            allow_redirects=True
        )

        if r.status_code != 200:

            print(
                f"❌ CATEGORY PAGE ERROR: "
                f"{r.status_code}"
            )

            return []

        html = r.text

        # =========================
        # DEBUG
        # =========================

        debug_path = os.path.join(
            OUTPUT_DIR,
            "category_debug.html"
        )

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        with open(
            debug_path,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(
                html
            )

        # =========================
        # BEAUTIFULSOUP
        # =========================

        soup = BeautifulSoup(
            html,
            "html.parser"
        )

        # =========================
        # ИЩЕМ МЕНЮ
        # =========================

        menu = soup.select_one(
            "ul#menu-category-menu-marketplace-2"
        )

        if not menu:

            links = soup.select(
                "a[href*='/product-category/']"
            )

            if not links:

                print(
                    "❌ КАТЕГОРИИ НЕ НАЙДЕНЫ"
                )

                return []

        else:

            links = menu.select(
                "li.item-level-0 "
                "a[href*='/product-category/']"
            )

        # =========================
        # ФОРМИРУЕМ КАТЕГОРИИ
        # =========================

        cats = []
        seen = set()

        for a in links:

            href = (
                a.get(
                    "href",
                    ""
                )
                .strip()
            )

            if not href:
                continue

            if "/product-category/" not in href:
                continue

            if not href.startswith(
                "http"
            ):

                href = (
                    BASE.rstrip("/")
                    + "/"
                    + href.lstrip("/")
                )

            href = href.split("?")[0]
            href = href.split("#")[0]

            if href in seen:
                continue

            seen.add(
                href
            )

            name = clean(
                a.get_text(
                    " ",
                    strip=True
                )
            )

            if not name:
                continue

            cats.append({
                "name": name,
                "url": href
            })

        print(
            f"📂 Найдено категорий: "
            f"{len(cats)}"
        )

        return cats

    except Exception as e:

        print(
            f"❌ CATEGORY ERROR: {e}"
        )

        return []


# =========================
# PRODUCT
# =========================

def parse_product(url):

    soup = get_soup(
        url,
        referer=BASE + "/"
    )

    # =========================
    # TITLE
    # =========================

    title = ""

    h1 = soup.select_one(
        "h1[itemprop='name']"
    )

    if not h1:

        h1 = soup.select_one(
            "h1.product_title"
        )

    if not h1:

        h1 = soup.select_one(
            "h1"
        )

    if h1:

        title = clean(
            h1.get_text(
                " ",
                strip=True
            )
        )

    # =========================
    # SKU
    # =========================

    sku = ""

    sku_el = soup.select_one(
        ".wd-product-detail.wd-product-sku .wd-sku"
    )

    if not sku_el:

        sku_el = soup.select_one(
            ".sku"
        )

    if sku_el:

        sku = clean(
            sku_el.get_text(
                " ",
                strip=True
            )
        )

    # =========================
    # PRICE
    # =========================

    price = ""

    price_el = soup.select_one(
        ".price .woocommerce-Price-amount"
    )

    if not price_el:

        price_el = soup.select_one(
            ".price"
        )

    if price_el:

        price = clean(
            price_el.get_text(
                " ",
                strip=True
            )
        )
    # =========================
    # STATUS
    # =========================
    
    status = ""
    
    cart_button = soup.select_one(
        "a.add-to-cart-loop"
    )
    
    if cart_button:
    
        status_el = cart_button.select_one(
            ".wd-action-text"
        )
    
        if status_el:
    
            status = clean(
                status_el.get_text(
                    " ",
                    strip=True
                )
            )
    
        else:
    
            status = clean(
                cart_button.get_text(
                    " ",
                    strip=True
                )
            )

    # =========================
    # DEBUG
    # =========================

    print(
        f"   📦 {title}"
        f" | SKU: {sku}"
        f" | PRICE: {price}"
        f" | STATUS: {status}"
    )

    return [
        sku,
        title,
        price,
        status,
        url
    ]

# =========================
# CATEGORY
# =========================

def parse_category(cat_url):

    result = []

    seen_skus = set()
    seen_urls = set()

    page = 1
    next_url = cat_url

    while True:

        print(
            f"📄 Страница {page}: "
            f"{next_url}"
        )

        soup = get_soup(
            next_url,
            referer=cat_url
        )

        # =========================
        # КАРТОЧКИ ТОВАРОВ
        # =========================

        product_cards = soup.select(
            "div.wd-product"
        )

        if not product_cards:

            print(
                f"⚠️ Товары не найдены "
                f"на странице {page}"
            )

            break

        added = 0

        for card in product_cards:

            # =========================
            # TITLE + URL
            # =========================

            product_link = card.select_one(
                "h3.wd-entities-title a[href]"
            )

            if not product_link:
                continue

            href = (
                product_link.get(
                    "href",
                    ""
                )
                .strip()
            )

            if not href:
                continue

            if not href.startswith("http"):

                href = (
                    BASE.rstrip("/")
                    + "/"
                    + href.lstrip("/")
                )

            href = href.split("?")[0]
            href = href.split("#")[0]

            href_key = (
                href
                .rstrip("/")
                .lower()
            )

            # =========================
            # TITLE
            # =========================

            title = clean(
                product_link.get_text(
                    " ",
                    strip=True
                )
            )

            if not title:
                continue

            # =========================
            # SKU
            # =========================

            sku = ""

            sku_element = card.select_one(
                "[data-product_sku]"
            )

            if sku_element:

                sku = clean(
                    sku_element.get(
                        "data-product_sku",
                        ""
                    )
                )

            sku_key = sku.lower()

            # =========================
            # ДЕДУБЛИКАЦИЯ
            # =========================

            # Если есть SKU —
            # считаем SKU главным идентификатором.

            if sku_key:

                if sku_key in seen_skus:
                    continue

            else:

                # Если SKU нет —
                # используем URL.

                if href_key in seen_urls:
                    continue

            # =========================
            # СОХРАНЯЕМ КЛЮЧИ
            # =========================

            if sku_key:
                seen_skus.add(sku_key)

            seen_urls.add(href_key)

            # =========================
            # PRICE
            # =========================

            price = ""

            price_element = card.select_one(
                ".price .woocommerce-Price-amount"
            )

            if price_element:

                price = clean(
                    price_element.get_text(
                        " ",
                        strip=True
                    )
                )

            # =========================
            # STATUS
            # =========================

            status = ""

            status_element = card.select_one(
                ".wd-add-btn .wd-action-text"
            )

            if status_element:

                status = clean(
                    status_element.get_text(
                        " ",
                        strip=True
                    )
                )

            if not status:

                button = card.select_one(
                    ".wd-add-btn a"
                )

                if button:

                    status = clean(
                        button.get_text(
                            " ",
                            strip=True
                        )
                    )

            # =========================
            # ДОБАВЛЯЕМ
            # =========================

            result.append([
                sku,
                title,
                price,
                status,
                href
            ])

            added += 1

        print(
            f"📦 Страница {page}: "
            f"{added} товаров"
        )

        # =========================
        # ПАГИНАЦИЯ
        # =========================

        load_more = soup.select_one(
            "a.wd-products-load-more[href]"
        )

        if not load_more:

            print(
                "🏁 Следующей страницы нет"
            )

            break

        next_href = (
            load_more.get(
                "href",
                ""
            )
            .strip()
        )

        if not next_href:

            print(
                "🏁 Ссылка следующей страницы пустая"
            )

            break

        if not next_href.startswith("http"):

            next_href = (
                BASE.rstrip("/")
                + "/"
                + next_href.lstrip("/")
            )

        # =========================
        # ЗАЩИТА ОТ ЦИКЛА
        # =========================

        if next_href == next_url:

            print(
                "🏁 Следующая страница "
                "совпадает с текущей"
            )

            break

        page += 1
        next_url = next_href

        if page > 100:

            print(
                "⚠️ Остановлено: "
                "слишком много страниц"
            )

            break

    print(
        f"📦 Всего уникальных товаров: "
        f"{len(result)}"
    )

    return result

# =========================
# MAIN
# =========================

def run_parser():

    if is_locked():
        return

    set_lock(True)

    try:

        save_status(
            True,
            0,
            USER,
            FILE_PATH
        )

        # =========================
        # LOGIN
        # =========================

        if not login():

            print(
                "❌ Авторизация не выполнена"
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        print(
            "🔓 Авторизация подтверждена"
        )

        # =========================
        # EXCEL
        # =========================

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

        # =========================
        # ОБЩАЯ ДЕДУПЛИКАЦИЯ
        # ДЛЯ ВСЕХ КАТЕГОРИЙ
        # =========================

        global_seen_skus = set()
        global_seen_urls = set()

        # =========================
        # CATEGORIES
        # =========================

        cats = get_categories()

        print(
            f"📂 Категорий: "
            f"{len(cats)}"
        )

        if CATEGORY_LIMIT:

            cats = cats[
                :CATEGORY_LIMIT
            ]

        total = len(cats)

        if total == 0:

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return

        # =========================
        # PARSE
        # =========================

        for i, cat in enumerate(
            cats,
            1
        ):

            print(
                f"📂 [{i}/{total}] "
                f"{cat['name']}"
            )

            save_status(
                True,
                int(
                    i
                    / total
                    * 100
                ),
                USER,
                FILE_PATH
            )

            items = parse_category(
                cat["url"]
            )

            for (
                sku,
                title,
                price,
                status,
                url
            ) in items:

                if not title:
                    continue

                # =========================
                # ДЕДУБЛИКАЦИЯ ПЕРЕД EXCEL
                # =========================

                sku_key = (
                    clean(sku)
                    .lower()
                    if sku
                    else ""
                )

                url_key = (
                    url
                    .split("?")[0]
                    .split("#")[0]
                    .rstrip("/")
                    .lower()
                    if url
                    else ""
                )

                # Если SKU уже встречался —
                # НЕ записываем товар в Excel

                if sku_key:

                    if sku_key in global_seen_skus:
                        continue

                    global_seen_skus.add(
                        sku_key
                    )

                else:

                    # Если SKU отсутствует —
                    # проверяем URL

                    if url_key in global_seen_urls:
                        continue

                if url_key:
                    global_seen_urls.add(
                        url_key
                    )

                # =========================
                # EXCEL
                # =========================

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    url
                ])

            time.sleep(
                0.2
            )

        # =========================
        # SAVE
        # =========================

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        tmp = (
            FILE_PATH
            + ".tmp"
        )

        wb.save(
            tmp
        )

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

        print(
            "✅ Готово. "
            "Харьковская 4425-4426 Gold Top"
        )

    finally:

        set_lock(False)


# =========================
# START
# =========================

if __name__ == "__main__":

    run_parser()

