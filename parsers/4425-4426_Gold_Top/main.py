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

OUTPUT_DIR = os.path.abspath("output/4425-4426_Gold_Top")
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

        os.makedirs(OUTPUT_DIR, exist_ok=True)

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
# LOGIN
# =========================

def login():

    print("🔐 LOGIN...")

    login_page = BASE + "/my-account/"

    try:

        time.sleep(2)

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:153.0) "
                "Gecko/20100101 Firefox/153.0"
            ),
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;"
                "q=0.9,image/avif,image/webp,*/*;q=0.8"
            ),
            "Accept-Language": (
                "uk-UA,uk;q=0.9,ru;q=0.8,en-US;q=0.7,en;q=0.6"
            ),
            "Cache-Control": "no-cache",
            "Pragma": "no-cache",
            "Referer": BASE + "/",
        }

        r = session.get(
            login_page,
            headers=headers,
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
            f"🍪 COOKIES: {dict(session.cookies)}"
        )

        if r.status_code != 200:

            print(
                f"❌ LOGIN PAGE ERROR: {r.status_code}"
            )

            return False

        # ==========================================
        # ПРОВЕРЯЕМ, ЕСТЬ ЛИ ФОРМА В ОТВЕТЕ
        # ==========================================

        if "woocommerce-login-nonce" in r.text:

            print("✅ woocommerce-login-nonce FOUND")

        else:

            print("❌ woocommerce-login-nonce NOT FOUND")

        if "woocommerce-form-login" in r.text:

            print("✅ woocommerce-form-login FOUND")

        else:

            print("❌ woocommerce-form-login NOT FOUND")

        if 'name="username"' in r.text:

            print("✅ username INPUT FOUND")

        else:

            print("❌ username INPUT NOT FOUND")

        if 'name="password"' in r.text:

            print("✅ password INPUT FOUND")

        else:

            print("❌ password INPUT NOT FOUND")

        # ==========================================
        # BEAUTIFULSOUP
        # ==========================================

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        forms = soup.find_all("form")

        print(
            f"🔎 FORMS FOUND: {len(forms)}"
        )

        for i, form in enumerate(forms, 1):

            print(
                f"FORM {i}: "
                f"class={form.get('class')} "
                f"action={form.get('action')}"
            )

        # ==========================================
        # ИЩЕМ ТОЧНО ТУ ФОРМУ,
        # КОТОРУЮ ТЫ ПРИСЛАЛА
        # ==========================================

        form = soup.select_one(
            "form.login.woocommerce-form.woocommerce-form-login"
        )

        if not form:

            print(
                "❌ LOGIN FORM NOT FOUND"
            )

            # Проверяем вообще наличие login form
            form = soup.select_one(
                "form.woocommerce-form-login"
            )

        if not form:

            print(
                "❌ LOGIN FORM STILL NOT FOUND"
            )

            print(
                "🔍 ПОИСК ПО ТЕКСТУ: ",
                "Увійти" in r.text
            )

            print(
                "🔍 ПОИСК USERNAME: ",
                'name="username"' in r.text
            )

            print(
                "🔍 ПОИСК PASSWORD: ",
                'name="password"' in r.text
            )

            # Сохраняем именно ответ Railway
            debug_file = os.path.join(
                OUTPUT_DIR,
                "login_debug.html"
            )

            os.makedirs(
                OUTPUT_DIR,
                exist_ok=True
            )

            with open(
                debug_file,
                "w",
                encoding="utf-8"
            ) as f:

                f.write(r.text)

            print(
                f"📄 DEBUG HTML: {debug_file}"
            )

            return False

        print(
            "✅ LOGIN FORM FOUND"
        )

        # ==========================================
        # ACTION
        # ==========================================

        action = form.get("action")

        print(
            "📌 FORM ACTION:",
            action
        )

        if not action:

            action = login_page

        # ==========================================
        # NONCE
        # ==========================================

        nonce = form.select_one(
            "input[name='woocommerce-login-nonce']"
        )

        if not nonce:

            print(
                "❌ LOGIN NONCE NOT FOUND"
            )

            return False

        nonce_value = nonce.get(
            "value",
            ""
        ).strip()

        print(
            "🔑 NONCE:",
            nonce_value
        )

        if not nonce_value:

            print(
                "❌ LOGIN NONCE EMPTY"
            )

            return False

        # ==========================================
        # REFERER
        # ==========================================

        referer_input = form.select_one(
            "input[name='_wp_http_referer']"
        )

        wp_referer = "/my-account/"

        if referer_input:

            wp_referer = referer_input.get(
                "value",
                "/my-account/"
            )

        print(
            "📌 WP REFERER:",
            wp_referer
        )

        # ==========================================
        # LOGIN POST
        # ==========================================

        payload = {
            "username": EMAIL,
            "password": PASSWORD,
            "woocommerce-login-nonce": nonce_value,
            "_wp_http_referer": wp_referer,
            "login": "Увійти",
            "rememberme": "forever"
        }

        print(
            "🚀 LOGIN POST..."
        )

        time.sleep(1)

        post_headers = {
            "User-Agent": headers["User-Agent"],
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;"
                "q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": headers["Accept-Language"],
            "Referer": login_page,
            "Origin": BASE,
        }

        r = session.post(
            action,
            data=payload,
            headers=post_headers,
            timeout=30,
            allow_redirects=True
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
            f"🍪 COOKIES AFTER LOGIN: "
            f"{dict(session.cookies)}"
        )

        if r.status_code == 429:

            print(
                "❌ LOGIN POST: 429"
            )

            return False

        # ==========================================
        # ПРОВЕРКА ПО КУКИ
        # ==========================================

        print(
            "🔎 CHECK LOGIN..."
        )

        check = session.get(
            login_page,
            headers={
                "User-Agent": headers["User-Agent"],
                "Accept": headers["Accept"],
                "Accept-Language": headers["Accept-Language"],
                "Referer": action,
                "Cache-Control": "no-cache",
            },
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔎 LOGIN CHECK: "
            f"{check.status_code} {check.url}"
        )

        print(
            f"📄 CHECK HTML LENGTH: "
            f"{len(check.text)}"
        )

        check_soup = BeautifulSoup(
            check.text,
            "html.parser"
        )

        # Если форма входа всё ещё присутствует,
        # значит авторизация не произошла.
        login_form = check_soup.select_one(
            "form.woocommerce-form-login"
        )

        if login_form:

            print(
                "❌ LOGIN FAIL: "
                "LOGIN FORM STILL PRESENT"
            )

            return False

        # Ищем признаки выхода
        logout = (
            check_soup.select_one(
                "a[href*='customer-logout']"
            )
            or
            check_soup.select_one(
                "a[href*='logout']"
            )
        )

        if logout:

            print(
                "✅ LOGIN OK"
            )

            return True

        text = check.text.lower()

        if (
            "вийти" in text
            or
            "вихід" in text
            or
            "выйти" in text
            or
            "logout" in text
        ):

            print(
                "✅ LOGIN OK"
            )

            return True

        print(
            "❌ LOGIN NOT CONFIRMED"
        )

        debug_file = os.path.join(
            OUTPUT_DIR,
            "account_debug.html"
        )

        with open(
            debug_file,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(check.text)

        print(
            f"📄 DEBUG HTML: {debug_file}"
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

    soup = get_soup(BASE)

    cats = []
    seen = set()

    for a in soup.select(
        "#d_category_menu_list a.link-level-1"
    ):

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

        name = clean(
            a.get_text()
        )

        # Убираем количество товаров
        name = re.sub(
            r"\d+\s*$",
            "",
            name
        ).strip()

        cats.append({
            "name": name,
            "url": href
        })

    return cats


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
                f"{cat_url}?page={page}"
            )

        soup = get_soup(url)

        products = soup.select(
            "div.product-name a"
        )

        if not products:
            break

        added = 0

        for a in products:

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:
                continue

            if "/image/" in href:
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

            result.append(
                parse_product(href)
            )

            added += 1

            time.sleep(0.05)

        if added == 0:
            break

        page += 1

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

    span = soup.select_one(
        "div.mr-4.p-1.text-secondary "
        "span.text-danger"
    )

    if span:

        sku = clean(
            span.get_text()
        )

    # =========================
    # PRICE
    # =========================

    price = ""

    p = soup.select_one(
        ".h2.m-0.text-nowrap"
    )

    if p:

        price = clean(
            p.get_text()
        )

    # =========================
    # STATUS
    # =========================

    status = ""

    alert = soup.select_one(
        ".alert"
    )

    if alert:

        status = clean(
            alert.get_text()
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


if __name__ == "__main__":

    run_parser()
