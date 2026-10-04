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

print("🔥 DELLTA LIFE PARSER")

BASE = "https://b2b.delltalife.com"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 2
# CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"

OUTPUT_DIR = os.path.abspath("output/Dellta")
FILE_PATH = os.path.join(OUTPUT_DIR, "Dellta_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/140.0.0.0 Safari/537.36"
}

session = requests.Session()
session.headers.update(HEADERS)

# =========================
# LOGIN
# =========================
def login():

    login_action = (
        BASE
        + "/themes/default/ajax/login.php"
    )

    payload = {
        "email_auth": EMAIL,
        "pass_auth": PASSWORD
    }

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/140.0.0.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
        "X-Requested-With": "XMLHttpRequest",
        "Origin": BASE,
        "Referer": BASE + "/",
        "Connection": "keep-alive"
    }

    print("🔐 Отправляем запрос авторизации Dellta...")

    try:

        r = session.post(
            login_action,
            data=payload,
            headers=headers,
            timeout=(15, 60)
        )

        print(
            f"🔐 LOGIN HTTP: {r.status_code}"
        )

        print(
            f"🔐 LOGIN RESPONSE: {r.text[:500]}"
        )

        if r.status_code == 429:

            print(
                "❌ Dellta вернул 429 на login.php"
            )

            return False

        r.raise_for_status()

    except requests.exceptions.RequestException as e:

        print(
            f"❌ LOGIN ERROR: {e}"
        )

        return False

    print("✅ LOGIN REQUEST OK")

    return True


# =========================
# STATUS
# =========================
def save_status(running=False, progress=0, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path
    }

    tmp = STATUS_PATH + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )

    os.replace(tmp, STATUS_PATH)


# =========================
# HTTP
# =========================
def get_soup(url):

    while True:

        try:

            r = session.get(
                url,
                timeout=(15, 60)
            )

            r.raise_for_status()

            return BeautifulSoup(
                r.text,
                "html.parser"
            )

        except requests.exceptions.Timeout:

            time.sleep(5)

        except requests.exceptions.RequestException:

            time.sleep(5)


# =========================
# CLEAN
# =========================
def clean(text):

    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        text
    ).strip()


# =========================
# CATEGORIES
# =========================
def get_categories():

    try:

        r = session.get(
            BASE,
            timeout=(15, 60)
        )

        r.raise_for_status()

    except requests.exceptions.RequestException:

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
        return categories

    for a in container.select(
        "a.COMitem"
    ):

        href = a.get("href")

        if not href:
            continue

        if href.startswith("/"):

            href = BASE.rstrip("/") + href

        elif not href.startswith("http"):

            href = BASE.rstrip("/") + "/" + href.lstrip("/")

        if href not in categories:

            categories.append(href)

    return categories


# =========================
# LAST PAGE DETECTION
# =========================
def get_last_page(soup):

    pages = []

    for a in soup.select(
        ".pagination .page-link[pn]"
    ):

        pn = a.get("pn")

        if pn and pn.isdigit():

            pages.append(
                int(pn)
            )

    return max(pages) if pages else 1


# =========================
# PARSE PRODUCT CARD
# =========================
def parse_product_card(card):

    title = ""
    sku = ""
    status = ""
    price = ""
    url = ""

    # =========================
    # SKU
    # =========================
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

    # =========================
    # TITLE + URL
    # =========================
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

                url = BASE.rstrip("/") + href

            elif href.startswith("http"):

                url = href

            else:

                url = BASE.rstrip("/") + "/" + href.lstrip("/")

    # =========================
    # FALLBACK URL
    # =========================
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

                url = BASE.rstrip("/") + href

            elif href.startswith("http"):

                url = href

    # =========================
    # STATUS
    # =========================
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

        # Более универсальный fallback
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

    # =========================
    # DEALER PRICE
    # =========================
    price_el = card.select_one(
        "td.td_3 .price-table tr.line-1 span.active"
    )

    if price_el:

        price = clean(
            price_el.get_text(
                " ",
                strip=True
            )
        )

    # =========================
    # FALLBACK PRICE
    # =========================
    if not price:

        price_el = card.select_one(
            "td.td_3 tr.line-1 span.active"
        )

        if price_el:

            price = clean(
                price_el.get_text(
                    " ",
                    strip=True
                )
            )

    # =========================
    # RETURN
    # =========================
    return [
        sku,
        title,
        price,
        status,
        url
    ]


# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    first_page = get_soup(
        cat_url
    )

    last_page = get_last_page(
        first_page
    )

    for page in range(
        1,
        last_page + 1
    ):

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

        # =========================
        # PRODUCT CARDS
        # =========================
        cards = soup.select(
            "tr.itemPosition"
        )

        for card in cards:

            item = parse_product_card(
                card
            )

            sku, title, price, status, url = item

            # Без названия товар пропускаем
            if not title:
                continue

            all_items.append(
                item
            )

    return all_items


# =========================
# MAIN
# =========================
def run_parser():

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

        save_status(
            False,
            0,
            USER,
            ""
        )

        print("❌ LOGIN FAILED")

        return

    # =========================
    # EXCEL
    # =========================
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

    # =========================
    # CATEGORIES
    # =========================
    cats = get_categories()

    if not cats:

        print("❌ CATEGORIES NOT FOUND")

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

    total = len(cats)

    # =========================
    # PARSE
    # =========================
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

        items = parse_category(
            cat
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

        time.sleep(0.3)

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

    print("DONE")


# =========================
# START
# =========================
if __name__ == "__main__":

    run_parser()
