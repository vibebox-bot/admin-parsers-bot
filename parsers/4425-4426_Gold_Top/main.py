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

CATEGORY_LIMIT = 1
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

            f.write(
                str(time.time())
            )

    else:

        if os.path.exists(LOCK_FILE):

            os.remove(LOCK_FILE)


# =========================
# COOKIES
# =========================

def print_cookies():

    print("🍪 COOKIES:")

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


def login():

    print("🔐 LOGIN...")

    login_page = BASE + "/my-account/"

    try:

        # ==========================================
        # LOGIN PAGE
        # ==========================================

        time.sleep(2)

        login_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;"
                "q=0.9,image/avif,image/webp,image/apng,*/*;"
                "q=0.8"
            ),
            "Accept-Language": HEADERS["Accept-Language"],
            # ВАЖНО: НЕ используем br
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

        print(
            f"🌐 LOGIN PAGE: {r.status_code} {r.url}"
        )

        print(
            f"📄 HTML LENGTH: {len(r.text)}"
        )

        print(
            f"📦 CONTENT-TYPE: "
            f"{r.headers.get('Content-Type')}"
        )

        print(
            f"📦 CONTENT-ENCODING: "
            f"{r.headers.get('Content-Encoding')}"
        )

        if r.status_code != 200:

            print(
                f"❌ LOGIN PAGE ERROR: "
                f"{r.status_code}"
            )

            return False

        # ==========================================
        # HTML
        # ==========================================

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        # ==========================================
        # LOGIN FORM
        # ==========================================

        form = soup.select_one(
            "form.woocommerce-form-login"
        )

        if not form:

            form = soup.select_one(
                "form.woocommerce-form.woocommerce-form-login"
            )

        if not form:

            print("❌ LOGIN FORM NOT FOUND")

            print(
                f"🔎 FORMS FOUND: "
                f"{len(soup.find_all('form'))}"
            )

            # Дополнительный поиск
            forms = soup.find_all("form")

            for i, f in enumerate(forms, 1):

                print(
                    f"FORM {i}: "
                    f"class={f.get('class')} "
                    f"action={f.get('action')}"
                )

            with open(
                os.path.join(
                    OUTPUT_DIR,
                    "login_debug.html"
                ),
                "w",
                encoding="utf-8"
            ) as f:

                f.write(r.text)

            return False

        print("✅ LOGIN FORM FOUND")

        # ==========================================
        # NONCE
        # ==========================================

        nonce = form.select_one(
            "input[name='woocommerce-login-nonce']"
        )

        if not nonce:

            print(
                "❌ woocommerce-login-nonce "
                "NOT FOUND"
            )

            return False

        nonce_value = (
            nonce.get("value", "")
            .strip()
        )

        if not nonce_value:

            print(
                "❌ LOGIN NONCE EMPTY"
            )

            return False

        print(
            "🔑 NONCE:",
            nonce_value[:10] + "..."
        )

        # ==========================================
        # FORM ACTION
        # ==========================================

        action = form.get("action")

        if not action:

            action = (
                login_page
                + "?action=login"
            )

        if not action.startswith("http"):

            action = (
                BASE.rstrip("/")
                + "/"
                + action.lstrip("/")
            )

        print(
            "🚀 POST URL:",
            action
        )

        # ==========================================
        # PAYLOAD
        # ==========================================

        payload = {
            "username": EMAIL,
            "password": PASSWORD,
            "woocommerce-login-nonce": nonce_value,
            "_wp_http_referer": "/my-account/",
            "login": "Увійти",
            "rememberme": "forever"
        }

        # ==========================================
        # LOGIN POST
        # ==========================================

        time.sleep(1)

        print("📤 LOGIN POST...")

        post_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": HEADERS["Accept-Language"],
            # ВАЖНО: опять без br
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

        print(
            f"🌐 LOGIN POST: "
            f"{r.status_code} {r.url}"
        )

        print(
            f"📄 POST HTML LENGTH: "
            f"{len(r.text)}"
        )

        print(
            f"📦 POST CONTENT-TYPE: "
            f"{r.headers.get('Content-Type')}"
        )

        print(
            f"📦 POST CONTENT-ENCODING: "
            f"{r.headers.get('Content-Encoding')}"
        )

        if r.status_code == 429:

            print(
                "❌ LOGIN POST: 429"
            )

            return False

        # ==========================================
        # НЕ ИСПОЛЬЗУЕМ cookies.get_dict()
        # ==========================================
        # У сайта могут быть два cookie
        # с одинаковым именем wordpress_sec_*.
        # Поэтому просто смотрим количество cookies.

        print(
            f"🍪 COOKIES COUNT: "
            f"{len(session.cookies)}"
        )

        # ==========================================
        # ПРОВЕРКА ОТВЕТА ПОСЛЕ LOGIN
        # ==========================================

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

        # ==========================================
        # УСПЕШНЫЙ LOGIN
        # ==========================================

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

        # ==========================================
        # ДОПОЛНИТЕЛЬНАЯ ПРОВЕРКА
        # ==========================================

        print(
            "🔎 LOGIN MARKER NOT FOUND"
        )

        # Если POST вернул обратно страницу
        # входа — проверяем наличие формы.
        login_form_after = response_soup.select_one(
            "form.woocommerce-form-login"
        )

        if login_form_after:

            print(
                "❌ LOGIN FAIL: "
                "LOGIN FORM STILL PRESENT"
            )

        else:

            print(
                "⚠️ LOGIN FORM NOT FOUND "
                "AFTER POST"
            )

        # ==========================================
        # СОХРАНЯЕМ ОТВЕТ ДЛЯ DEBUG
        # ==========================================

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

            f.write(r.text)

        print(
            f"📄 DEBUG POST HTML: "
            f"{debug_path}"
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

def get_soup(url):

    for attempt in range(3):

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
                    f"⚠️ 429: {url}"
                )

                time.sleep(
                    5 * (attempt + 1)
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

    print("📂 ПОИСК КАТЕГОРИЙ...")

    try:

        category_headers = {
            "User-Agent": HEADERS["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,"
                "application/xml;q=0.9,image/avif,"
                "image/webp,image/apng,*/*;q=0.8"
            ),
            "Accept-Language": HEADERS["Accept-Language"],
            "Referer": BASE + "/my-account/",
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
        }

        r = session.get(
            BASE + "/",
            headers=category_headers,
            params={
                "_": str(int(time.time() * 1000))
            },
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🌐 CATEGORY PAGE: {r.status_code} {r.url}"
        )

        print(
            f"📄 CATEGORY HTML LENGTH: {len(r.text)}"
        )

        print(
            f"📦 CONTENT-TYPE: "
            f"{r.headers.get('Content-Type')}"
        )

        print(
            f"🍪 COOKIES COUNT: "
            f"{len(session.cookies)}"
        )

        if r.status_code != 200:

            print(
                f"❌ CATEGORY PAGE ERROR: "
                f"{r.status_code}"
            )

            return []

        html = r.text

        # =================================================
        # ПРОВЕРКА ИМЕННО ТОГО БЛОКА,
        # КОТОРЫЙ ТЫ ДАЛА
        # =================================================

        if "menu-category-menu-marketplace-2" in html:

            print(
                "✅ БЛОК КАТЕГОРИЙ НАЙДЕН"
            )

        else:

            print(
                "❌ БЛОК КАТЕГОРИЙ НЕ НАЙДЕН"
            )

            if "/product-category/" in html:

                print(
                    "⚠️ product-category ЕСТЬ"
                )

            else:

                print(
                    "❌ product-category НЕТ"
                )

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

                f.write(html)

            print(
                f"📄 DEBUG HTML: {debug_path}"
            )

            return []

        # =================================================
        # PARSE
        # =================================================

        soup = BeautifulSoup(
            html,
            "html.parser"
        )

        menu = soup.select_one(
            "#menu-category-menu-marketplace-2"
        )

        if not menu:

            print(
                "❌ MENU ELEMENT NOT FOUND"
            )

            return []

        cats = []
        seen = set()

        # =================================================
        # КАТЕГОРИИ
        # =================================================

        for a in menu.select(
            "a.woodmart-nav-link[href]"
        ):

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:
                continue

            if "/product-category/" not in href:
                continue

            if not href.startswith("http"):

                href = (
                    BASE.rstrip("/")
                    + "/"
                    + href.lstrip("/")
                )

            href = href.split("?")[0].rstrip("/")

            if href in seen:
                continue

            name_el = a.select_one(
                ".nav-link-text"
            )

            if name_el:

                name = clean(
                    name_el.get_text()
                )

            else:

                name = clean(
                    a.get_text()
                )

            if not name:
                continue

            seen.add(href)

            cats.append({
                "name": name,
                "url": href + "/"
            })

            print(
                f"📂 {name}"
            )

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
# CATEGORY
# =========================

def parse_category(cat_url):

    result = []
    seen = set()

    page = 1

    while True:

        if page == 1:

            url = cat_url

        else:

            url = (
                f"{cat_url.rstrip('/')}"
                f"/page/{page}/"
            )

        print(
            f"📄 CATEGORY PAGE: "
            f"{url}"
        )

        soup = get_soup(url)

        # =========================
        # КАРТОЧКИ ТОВАРОВ
        # =========================

        products = soup.select(
            "h3.wd-entities-title "
            "a[href]"
        )

        if not products:

            print(
                f"⚠️ Товары не найдены "
                f"на странице {page}"
            )

            break

        added = 0

        for a in products:

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:

                continue

            if not href.startswith("http"):

                href = (
                    BASE.rstrip("/")
                    + "/"
                    + href.lstrip("/")
                )

            if href in seen:

                continue

            seen.add(href)

            product = parse_product(
                href
            )

            result.append(
                product
            )

            added += 1

            time.sleep(0.05)

        print(
            f"📦 Страница {page}: "
            f"{added} товаров"
        )

        if added == 0:

            break

        page += 1

    print(
        f"📦 Всего товаров в категории: "
        f"{len(result)}"
    )

    return result


# =========================
# PRODUCT
# =========================

def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================

    title = ""

    h1 = soup.select_one(
        "h1[itemprop='name']"
    )

    if h1:

        title = clean(
            h1.get_text()
        )

    # =========================
    # SKU
    # =========================

    sku = ""

    sku_el = soup.select_one(
        ".wd-product-detail.wd-product-sku "
        ".wd-sku"
    )

    if sku_el:

        sku = clean(
            sku_el.get_text()
        )

    # =========================
    # PRICE
    # =========================

    price = ""

    price_el = soup.select_one(
        ".price "
        ".woocommerce-Price-amount"
    )

    if price_el:

        price = clean(
            price_el.get_text()
        )

    # =========================
    # STATUS
    # =========================
    #
    # НИЧЕГО НЕ ПРИДУМЫВАЕМ.
    #
    # Берём РОВНО текст кнопки.
    #
    # Например:
    # "Додати в кошик"
    #
    # так и попадёт в Excel.
    # =========================

    status = ""

    cart_button = soup.select_one(
        "a.add-to-cart-loop"
    )

    if cart_button:

        status = clean(
            cart_button.get_text(
                " ",
                strip=True
            )
        )

    # =========================
    # DEBUG PRODUCT
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
        # CATEGORIES
        # =========================

        cats = get_categories()

        print(
            f"📂 Категорий: {len(cats)}"
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
                    i / total * 100
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

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    url
                ])

            time.sleep(0.2)

        # =========================
        # SAVE
        # =========================

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

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
