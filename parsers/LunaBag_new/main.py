import os
import json
import re
import time
import requests
from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 LunaBag_new")


BASE = "https://lunabag.com.ua"


# =========================
# ⚙️ SWITCH
# =========================

#CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None


OUTPUT_DIR = os.path.abspath("output/LunaBag_new")
FILE_PATH = os.path.join(OUTPUT_DIR, "LunaBag_new_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,en;q=0.8"
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

            os.remove(
                LOCK_FILE
            )


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
# URL
# =========================

def normalize_url(url):

    if not url:
        return ""

    url = url.strip()

    url = urljoin(
        BASE,
        url
    )

    parsed = urlparse(url)

    parsed = parsed._replace(
        fragment=""
    )

    return urlunparse(
        parsed
    )


def page_url(base_url, page):

    if page == 1:
        return base_url

    parsed = urlparse(base_url)

    query = parse_qs(
        parsed.query,
        keep_blank_values=True
    )

    query["page"] = [str(page)]

    new_query = urlencode(
        query,
        doseq=True
    )

    return urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            new_query,
            ""
        )
    )


# =========================
# HTTP
# =========================

def get_soup(url):

    for attempt in range(3):

        try:

            r = session.get(
                url,
                timeout=40
            )

            if r.status_code == 200:

                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

        except:

            pass

        time.sleep(2)

    return BeautifulSoup(
        "",
        "html.parser"
    )


# =========================
# CURRENCY USD
# =========================

def set_currency_usd():

    try:

        r = session.get(
            BASE,
            timeout=30
        )

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        form = soup.select_one(
            "#form-currency"
        )

        if not form:
            return

        action = form.get(
            "action",
            ""
        ).strip()

        if not action:

            action = (
                BASE
                + "/index.php?route=common/currency/currency"
            )

        action = normalize_url(
            action
        )

        payload = {}

        for inp in form.select(
            "input"
        ):

            name = inp.get("name")

            if not name:
                continue

            payload[name] = inp.get(
                "value",
                ""
            )

        payload["code"] = "USD"

        if "redirect" not in payload:

            payload["redirect"] = BASE

        session.post(
            action,
            data=payload,
            headers={
                "Referer": BASE
            },
            timeout=30,
            allow_redirects=True
        )

    except:

        pass


# =========================
# CATEGORIES
# =========================

def get_categories():

    soup = get_soup(
        BASE
    )

    categories = []

    seen = set()

    menu = soup.select_one(
        "nav.ds-menu-catalog-inner"
    )

    if not menu:

        print(
            "⚠ Categories not found"
        )

        return categories


    def add_category(url):

        url = normalize_url(
            url
        )

        if not url:
            return

        parsed = urlparse(url)

        if parsed.netloc.lower() != urlparse(BASE).netloc.lower():
            return

        if url.startswith(
            "javascript:"
        ):
            return

        if "route=product/product" in url:
            return

        if "route=account/" in url:
            return

        if "route=checkout/" in url:
            return

        if "route=information/" in url:
            return

        if url not in seen:

            seen.add(
                url
            )

            categories.append(
                url
            )


    for a in menu.select(
        "a[href]"
    ):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:
            continue

        if href == "#":
            continue

        if href.startswith(
            "javascript:"
        ):
            continue

        add_category(
            href
        )


    print(
        f"📂 Categories: "
        f"{len(categories)}"
    )

    return categories


# =========================
# PAGINATION
# =========================

def get_last_page(soup):

    pages = [1]

    for a in soup.select(
        "ul.pagination a[href]"
    ):

        href = a.get(
            "href",
            ""
        )

        if not href:
            continue

        parsed = urlparse(
            href
        )

        query = parse_qs(
            parsed.query
        )

        if "page" not in query:
            continue

        for value in query["page"]:

            try:

                page = int(
                    value
                )

                if page > 0:

                    pages.append(
                        page
                    )

            except:

                pass

    return max(
        pages
    )


# =========================
# PRODUCT PAGE STATUS
# =========================

def get_product_status(url):

    if not url:
        return ""

    try:

        soup = get_soup(
            url
        )

        status_el = soup.select_one(
            ".ds-product-main-stock"
        )

        if status_el:

            return clean(
                status_el.get_text(
                    " ",
                    strip=True
                )
            )

    except:

        pass

    return ""


# =========================
# PRODUCT CARD
# =========================

def parse_card(card):

    result = {

        "sku": "",

        "title": "",

        "price": "",

        "status": "",

        "url": ""
    }


    # =========================
    # TITLE + URL
    # =========================

    title_el = card.select_one(
        ".ds-module-title"
    )

    if title_el:

        result["title"] = clean(
            title_el.get_text(
                " ",
                strip=True
            )
        )

        href = title_el.get(
            "href",
            ""
        ).strip()

        result["url"] = normalize_url(
            href
        )


    # =========================
    # SKU
    # =========================

    code_el = card.select_one(
        ".ds-module-code"
    )

    if code_el:

        result["sku"] = clean(
            code_el.get_text(
                " ",
                strip=True
            )
        )

        result["sku"] = re.sub(
            r"^\s*Код\s+товару\s*:\s*",
            "",
            result["sku"],
            flags=re.I
        ).strip()


    # =========================
    # STATUS
    # =========================

    status_el = card.select_one(
        ".ds-module-stock"
    )

    if status_el:

        result["status"] = clean(
            status_el.get_text(
                " ",
                strip=True
            )
        )


    # =========================
    # PRICE
    # =========================

    price_el = card.select_one(
        ".ds-price-new"
    )

    if price_el:

        result["price"] = clean(
            price_el.get_text(
                " ",
                strip=True
            )
        )


    # =========================
    # FALLBACK STATUS
    # =========================

    if not result["status"] and result["url"]:

        result["status"] = get_product_status(
            result["url"]
        )


    return result


# =========================
# PARSE CATEGORY
# =========================

def parse_category(
    cat_url
):

    all_items = []

    first_page = get_soup(
        cat_url
    )

    if not first_page:

        return all_items

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

            url = page_url(
                cat_url,
                page
            )

            soup = get_soup(
                url
            )


        cards = soup.select(
            "div.ds-module-item.product-layout"
        )


        for card in cards:

            product = parse_card(
                card
            )

            if not product["title"]:

                continue

            all_items.append(
                [
                    product["sku"],
                    product["title"],
                    product["price"],
                    product["status"],
                    product["url"]
                ]
            )


        time.sleep(
            0.2
        )


    return all_items


# =========================
# MAIN
# =========================

def run_parser():

    if is_locked():

        print(
            "🔒 LunaBag already running"
        )

        return


    set_lock(
        True
    )


    try:

        # =========================
        # START
        # =========================

        save_status(
            True,
            0,
            USER,
            FILE_PATH
        )


        # =========================
        # CURRENCY
        # =========================

        set_currency_usd()


        # =========================
        # EXCEL
        # =========================

        wb = Workbook()

        ws = wb.active

        ws.title = "LunaBag"

        ws.append(
            [
                "Артикул",
                "Название",
                "Цена $",
                "Статус",
                "Ссылка"
            ]
        )


        # =========================
        # CATEGORIES
        # =========================

        categories = get_categories()


        if CATEGORY_LIMIT:

            categories = categories[
                :CATEGORY_LIMIT
            ]


        total = len(
            categories
        )


        if total == 0:

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return


        # =========================
        # DEDUP
        # =========================

        seen = set()


        # =========================
        # PARSE
        # =========================

        for i, cat in enumerate(
            categories,
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


            for (
                sku,
                title,
                price,
                status,
                url
            ) in items:


                # =========================
                # DEDUP
                # =========================

                if sku:

                    key = (
                        "sku",
                        sku
                    )

                elif url:

                    key = (
                        "url",
                        url
                    )

                else:

                    key = (
                        "product",
                        title,
                        price
                    )


                if key in seen:

                    continue


                seen.add(
                    key
                )


                # =========================
                # WRITE
                # =========================

                ws.append(
                    [
                        sku,
                        title,
                        price,
                        status,
                        url
                    ]
                )


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

        tmp = FILE_PATH + ".tmp"

        wb.save(
            tmp
        )

        os.replace(
            tmp,
            FILE_PATH
        )


        # =========================
        # FINISH
        # =========================

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )


        print(
            "✅ Готово LunaBag_new"
        )

    except Exception as e:

        print(
            "❌ PARSER ERROR:",
            e
        )

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        raise


    finally:

        set_lock(
            False
        )


# =========================
# START
# =========================

if __name__ == "__main__":

    run_parser()
