import os
import json
import re
import time
import requests
from datetime import datetime
from io import BytesIO

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

import sys


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 ARIZONE Sports")

BASE = "https://arizonesports.com.ua/uk"

GOOGLE_SHEET_ID = "1ROCGu-3W8PotpQSIgzHRDkXCiDeHU0_dzEjuRnkNHv0"

# =========================
# ⚙️ SWITCH
# =========================

# Для теста
#CATEGORY_LIMIT = 2
#PRODUCT_LIMIT = 50

# Для полного парсинга
CATEGORY_LIMIT = None
PRODUCT_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/ARIZONE_Sports")
FILE_PATH = os.path.join(OUTPUT_DIR, "ARIZONE_Sports_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7"
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

        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:

        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


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
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, STATUS_PATH)


# =========================
# HTTP
# =========================

def get_soup(url):

    for attempt in range(3):

        try:

            r = session.get(url, timeout=30)

            if r.status_code == 200:

                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

        except Exception:
            pass

        time.sleep(1)

    return BeautifulSoup("", "html.parser")


def clean(t):

    if not t:
        return ""

    t = str(t)

    t = t.replace("\xa0", " ")
    t = t.replace("\u200b", "")
    t = t.replace("\ufeff", "")

    return re.sub(r"\s+", " ", t).strip()


# =========================
# URL
# =========================

def absolute_url(url):

    if not url:
        return ""

    url = url.strip()

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("/"):
        return "https://arizonesports.com.ua" + url

    if url.startswith("http://") or url.startswith("https://"):
        return url

    return BASE + "/" + url.lstrip("/")


# =========================
# SKU NORMALIZATION
# =========================

def normalize_sku(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = value.replace("\u200b", "")
    value = value.replace("\ufeff", "")

    value = value.strip()

    if re.fullmatch(r"\d+\.0+", value):
        value = value.split(".")[0]

    value = re.sub(r"\s+", "", value)

    return value.casefold()


# =========================
# PRICE PARSING
# =========================

def parse_price(value):

    if value is None:
        return None

    if isinstance(value, (int, float)):

        try:
            return float(value)
        except:
            return None

    value = str(value).strip()

    if not value:
        return None

    value = value.replace("\xa0", " ")
    value = value.replace("$", "")
    value = value.replace("USD", "")
    value = value.replace("usd", "")
    value = value.strip()

    value = value.replace(" ", "")

    if "," in value and "." in value:

        if value.rfind(",") > value.rfind("."):

            value = value.replace(".", "")
            value = value.replace(",", ".")

        else:

            value = value.replace(",", "")

    elif "," in value:

        value = value.replace(",", ".")

    try:
        return float(value)

    except:
        return None


# =========================
# GOOGLE SHEET
# =========================

def load_google_prices():

    export_url = (
        "https://docs.google.com/spreadsheets/d/"
        + GOOGLE_SHEET_ID
        + "/export?format=xlsx"
    )

    try:

        response = session.get(
            export_url,
            timeout=60
        )

        if response.status_code != 200:
            return {}

        if len(response.content) < 1000:
            return {}

        try:

            wb = load_workbook(
                BytesIO(response.content),
                data_only=True,
                read_only=True
            )

        except Exception:
            return {}

        price_map = {}

        for sheet_name in wb.sheetnames:

            ws = wb[sheet_name]

            rows = ws.iter_rows(
                values_only=True
            )

            try:
                header_row = next(rows)
            except StopIteration:
                continue

            header_row = list(header_row)

            article_col = None
            price_col = None

            for index, value in enumerate(header_row):

                header = clean(value)

                if header == "Артикул":
                    article_col = index

                elif header == "Оптова ціна, $":
                    price_col = index

            if article_col is None:
                continue

            if price_col is None:
                continue

            for row in rows:

                row = list(row)

                if article_col >= len(row):
                    continue

                if price_col >= len(row):
                    continue

                raw_sku = row[article_col]
                raw_price = row[price_col]

                sku = normalize_sku(raw_sku)
                price = parse_price(raw_price)

                if not sku:
                    continue

                if price is None:
                    continue

                if sku in price_map:
                    continue

                price_map[sku] = price

        wb.close()

        return price_map

    except Exception:
        return {}


# =========================
# CATEGORIES
# =========================

def get_categories():

    soup = get_soup(BASE)

    categories = []
    seen = set()

    tree = soup.select_one(
        "#categories_block_left ul.tree"
    )

    if not tree:
        return categories

    for a in tree.select("a[href]"):

        href = a.get("href", "").strip()

        if not href:
            continue

        href = absolute_url(href)

        href = href.split("#")[0]

        if href in seen:
            continue

        seen.add(href)

        categories.append(href)

    return categories


# =========================
# LAST PAGE
# =========================

def get_last_page(soup):

    pages = [1]

    for a in soup.select(
        "#pagination_bottom a, "
        ".pagination a"
    ):

        href = a.get("href", "")

        m = re.search(
            r"page-(\d+)",
            href
        )

        if m:

            pages.append(
                int(m.group(1))
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
# CATEGORY PRODUCTS
# =========================

def get_category_products(cat_url):

    products = []
    seen = set()

    first_page = get_soup(cat_url)

    # ---------------------------------
    # SHOW ALL
    # ---------------------------------

    show_all_url = None

    form = first_page.select_one(
        "#pagination_bottom form.showall"
    )

    if form:

        action = form.get("action", "").strip()

        if action:

            params = []

            for inp in form.select(
                "input[type='hidden']"
            ):

                name = inp.get("name")
                value = inp.get("value", "")

                if name:
                    params.append(
                        f"{name}={value}"
                    )

            if params:

                separator = "?"

                if "?" in action:
                    separator = "&"

                show_all_url = (
                    action
                    + separator
                    + "&".join(params)
                )

    if show_all_url:

        soup = get_soup(
            show_all_url
        )

        product_links = extract_product_links(
            soup
        )

        for url in product_links:

            if url in seen:
                continue

            seen.add(url)
            products.append(url)

        if products:

            if (
                PRODUCT_LIMIT
                and len(products) >= PRODUCT_LIMIT
            ):
                return products[:PRODUCT_LIMIT]

            return products

    # ---------------------------------
    # PAGINATION
    # ---------------------------------

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

            separator = "?"

            if "?" in cat_url:
                separator = "&"

            url = (
                cat_url
                + separator
                + f"page={page}"
            )

            soup = get_soup(url)

        product_links = extract_product_links(
            soup
        )

        for url in product_links:

            if url in seen:
                continue

            seen.add(url)
            products.append(url)

            if (
                PRODUCT_LIMIT
                and len(products) >= PRODUCT_LIMIT
            ):

                return products

    return products


# =========================
# EXTRACT PRODUCT LINKS
# =========================

def extract_product_links(soup):

    result = []
    seen = set()

    selectors = [

        "ul.product_list li.ajax_block_product "
        "a.product-name",

        ".product-container a.product-name",

        "li.product-box.item "
        "a.product-image",

        "li.product-box.item "
        "h5.product-name a",

        "li.product-box.item "
        "a[href*='.html']",

        "#productscategory_list li.product-box "
        "a[href*='.html']"
    ]

    for selector in selectors:

        for a in soup.select(selector):

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:
                continue

            href = absolute_url(href)

            href = href.split("#")[0]

            if ".html" not in href:
                continue

            if href in seen:
                continue

            seen.add(href)
            result.append(href)

    return result


# =========================
# PRODUCT SKU
# =========================

def get_product_sku(soup):

    # 1. Schema
    for meta in soup.select(
        "meta[itemprop='sku']"
    ):

        value = clean(
            meta.get("content", "")
        )

        if value:
            return value

    # 2. itemprop
    sku_el = soup.select_one(
        "[itemprop='sku']"
    )

    if sku_el:

        value = sku_el.get(
            "content",
            ""
        )

        if not value:

            value = sku_el.get_text(
                " ",
                strip=True
            )

        value = clean(value)

        if value:
            return value

    # 3. Стандартные блоки
    selectors = [

        "#product_reference",
        ".product-reference",
        ".reference",
        ".product-info .reference"
    ]

    for selector in selectors:

        el = soup.select_one(selector)

        if not el:
            continue

        text = clean(
            el.get_text(
                " ",
                strip=True
            )
        )

        if not text:
            continue

        m = re.search(
            r"(?:Артикул|Код|Reference|SKU)"
            r"\s*[:\-]?\s*([A-Za-zА-Яа-я0-9_.\-]+)",
            text,
            re.I
        )

        if m:

            return clean(
                m.group(1)
            )

        return text

    # 4. По тексту страницы
    text = clean(
        soup.get_text(
            " ",
            strip=True
        )
    )

    patterns = [

        r"Артикул\s*[:\-]?\s*"
        r"([A-Za-zА-Яа-я0-9_.\-]+)",

        r"Код\s*товара\s*[:\-]?\s*"
        r"([A-Za-zА-Яа-я0-9_.\-]+)",

        r"Reference\s*[:\-]?\s*"
        r"([A-Za-zА-Яа-я0-9_.\-]+)",

        r"SKU\s*[:\-]?\s*"
        r"([A-Za-zА-Яа-я0-9_.\-]+)"
    ]

    for pattern in patterns:

        m = re.search(
            pattern,
            text,
            re.I
        )

        if m:

            return clean(
                m.group(1)
            )

    return ""


# =========================
# PRODUCT TITLE
# =========================

def get_product_title(soup):

    selectors = [

        "h1[itemprop='name']",

        "h1.product-title",

        "h1",

        ".product-name h1",

        ".product_name h1"
    ]

    for selector in selectors:

        el = soup.select_one(
            selector
        )

        if el:

            title = clean(
                el.get_text(
                    " ",
                    strip=True
                )
            )

            if title:
                return title

    return ""


# =========================
# PRODUCT STATUS
# =========================

def get_product_status(soup):

    availability = soup.select_one(
        "#availability_value"
    )

    if availability:

        text = clean(
            availability.get_text(
                " ",
                strip=True
            )
        )

        if text:
            return text

    availability = soup.select_one(
        "[itemprop='availability']"
    )

    if availability:

        value = (
            availability.get("href")
            or availability.get("content")
            or availability.get_text(
                " ",
                strip=True
            )
        )

        value = clean(value)

        if value:

            if "InStock" in value:
                return "В наявності"

            if "OutOfStock" in value:
                return "Немає в наявності"

            return value

    selectors = [

        ".availability",

        ".product-availability",

        "#availability"
    ]

    for selector in selectors:

        el = soup.select_one(
            selector
        )

        if not el:
            continue

        text = clean(
            el.get_text(
                " ",
                strip=True
            )
        )

        if text:
            return text

    return ""


# =========================
# PRODUCT COLORS
# =========================

def get_product_colors(soup):

    colors = []
    seen = set()

    for fieldset in soup.select(
        "fieldset.attribute_fieldset"
    ):

        label = fieldset.select_one(
            "label.attribute_label"
        )

        if not label:
            continue

        label_text = clean(
            label.get_text(
                " ",
                strip=True
            )
        ).casefold()

        if "колір" not in label_text:
            continue

        for li in fieldset.select(
            ".attribute_list ul > li"
        ):

            radio = li.select_one(
                "input.attribute_radio"
            )

            if not radio:
                continue

            color_id = clean(
                radio.get(
                    "value",
                    ""
                )
            )

            color_name = ""

            # Название цвета:
            # <span>Бежевий</span>
            for span in li.select(
                ":scope > span"
            ):

                if span.select_one("input"):
                    continue

                text = clean(
                    span.get_text(
                        " ",
                        strip=True
                    )
                )

                if text:
                    color_name = text
                    break

            if not color_name:
                continue

            key = (
                color_id,
                color_name.casefold()
            )

            if key in seen:
                continue

            seen.add(key)

            colors.append({
                "id": color_id,
                "name": color_name
            })

    if not colors:

        return [
            {
                "id": "",
                "name": ""
            }
        ]

    return colors


# =========================
# COLOR URL
# =========================

def make_color_url(
    product_url,
    color_id,
    color_name
):

    # Название цвета НЕ транслитерируем.
    # Если у сайта нет готовой ссылки,
    # оставляем обычную ссылку товара.

    return product_url


# =========================
# PARSE PRODUCT
# =========================

def parse_product(
    product_url,
    price_map
):

    soup = get_soup(
        product_url
    )

    title = get_product_title(
        soup
    )

    sku = get_product_sku(
        soup
    )

    status = get_product_status(
        soup
    )

    colors = get_product_colors(
        soup
    )

    normalized_sku = normalize_sku(
        sku
    )

    price = price_map.get(
        normalized_sku
    )

    rows = []

    for color in colors:

        color_name = color["name"]
        color_id = color["id"]

        color_url = make_color_url(
            product_url,
            color_id,
            color_name
        )

        rows.append(
            [
                sku,
                title,
                price,
                status,
                color_name,
                color_url
            ]
        )

    return rows


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

        # =================================
        # GOOGLE PRICES
        # =================================

        price_map = load_google_prices()

        if not price_map:

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # =================================
        # EXCEL
        # =================================

        wb = Workbook()

        ws = wb.active

        ws.title = "ARIZONE"

        ws.append(
            [
                "SKU",
                "TITLE",
                "PRICE USD",
                "STATUS",
                "COLOR",
                "URL"
            ]
        )

        # =================================
        # CATEGORIES
        # =================================

        cats = get_categories()

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

        # =================================
        # PRODUCTS
        # =================================

        seen_products = set()

        missing_prices = set()

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

            product_urls = get_category_products(
                cat
            )

            if PRODUCT_LIMIT:

                product_urls = product_urls[
                    :PRODUCT_LIMIT
                ]

            for product_url in product_urls:

                if product_url in seen_products:
                    continue

                seen_products.add(
                    product_url
                )

                rows = parse_product(
                    product_url,
                    price_map
                )

                for row in rows:

                    sku = row[0]
                    price = row[2]

                    if price is None:

                        missing_prices.add(
                            normalize_sku(sku)
                        )

                    ws.append(
                        row
                    )

                time.sleep(0.1)

        # =================================
        # SAVE
        # =================================

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
            "✅ Готово. ARIZONE Sports"
        )

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
