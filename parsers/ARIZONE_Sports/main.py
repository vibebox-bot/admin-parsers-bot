import os
import json
import re
import time
import requests
from io import BytesIO
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Самокаты ARIZONE Sports")

BASE = "https://arizonesports.com.ua/uk"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 2
#CATEGORY_LIMIT = None

PRODUCT_LIMIT = 20
#PRODUCT_LIMIT = None

# =========================
# GOOGLE SHEETS
# =========================

GOOGLE_SHEET_ID = "1ROCGu-3W8PotpQSIgzHRDkXCiDeHU0_dzEjuRnkNHv0"

# =========================
# OUTPUT
# =========================

OUTPUT_DIR = os.path.abspath("output/ARIZONE_Sports")
FILE_PATH = os.path.join(OUTPUT_DIR, "ARIZONE_Sports_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
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

    for _ in range(3):

        try:

            r = session.get(
                url,
                timeout=30
            )

            if r.status_code == 200:
                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

        except:
            pass

        time.sleep(1)

    return BeautifulSoup("", "html.parser")


def clean(t):

    return re.sub(
        r"\s+",
        " ",
        t
    ).strip() if t else ""


# =========================
# URL
# =========================

def absolute_url(url):

    if not url:
        return ""

    url = url.strip()

    if url.startswith("http://") or url.startswith("https://"):
        return url

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("/"):
        return "https://arizonesports.com.ua" + url

    return BASE.rstrip("/") + "/" + url.lstrip("/")


def clean_product_url(url):

    if not url:
        return ""

    # Цвета находятся после #
    url = url.split("#")[0]

    return url.strip()


# =========================
# GOOGLE SHEETS
# =========================

def normalize_header(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = value.replace("\u200b", "")

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip().casefold()


def normalize_sku(value):

    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", "")
    value = value.replace("\u200b", "")
    value = value.strip()

    # Excel может отдавать числовой артикул как 1918.0
    if re.fullmatch(r"\d+\.0+", value):
        value = value.split(".")[0]

    # Убираем пробелы внутри артикула
    value = re.sub(
        r"\s+",
        "",
        value
    )

    return value.casefold()


def parse_price(value):

    if value is None:
        return None

    if isinstance(value, (int, float)):
        return float(value)

    value = str(value).strip()

    if not value:
        return None

    value = value.replace("\xa0", " ")
    value = value.replace("$", "")
    value = value.replace("USD", "")
    value = value.strip()

    # Убираем пробелы-разделители тысяч
    value = value.replace(" ", "")

    if "," in value and "." in value:

        # Например 1.250,50
        if value.rfind(",") > value.rfind("."):

            value = value.replace(".", "")
            value = value.replace(",", ".")

        # Например 1,250.50
        else:

            value = value.replace(",", "")

    elif "," in value:

        value = value.replace(",", ".")

    try:

        return float(value)

    except:

        return None


def get_google_sheet_price_map():

    print("")
    print("=" * 70)
    print("📊 GOOGLE SHEETS")
    print("=" * 70)

    url = (
        "https://docs.google.com/spreadsheets/d/"
        + GOOGLE_SHEET_ID
        + "/export?format=xlsx"
    )

    print("📥 Загружаем всю таблицу...")

    try:

        r = session.get(
            url,
            timeout=90
        )

        print(
            f"HTTP: {r.status_code}"
        )

        print(
            f"Размер: {len(r.content):,} bytes"
        )

        r.raise_for_status()

    except Exception as e:

        print("❌ Ошибка загрузки Google Sheets:")
        print(e)

        return {}

    # XLSX — это ZIP, поэтому начинается с PK
    if not r.content.startswith(b"PK"):

        print("❌ Google не вернула XLSX.")

        try:
            print(
                r.text[:500]
            )
        except:
            pass

        return {}

    try:

        wb = load_workbook(
            BytesIO(r.content),
            data_only=True,
            read_only=True
        )

    except Exception as e:

        print("❌ Ошибка открытия Google Sheets:")
        print(e)

        return {}

    prices = {}

    print("")
    print(
        f"📚 Вкладок найдено: "
        f"{len(wb.sheetnames)}"
    )

    duplicate_count = 0

    # =========================
    # ВСЕ ВКЛАДКИ
    # =========================

    for sheet_name in wb.sheetnames:

        print("")
        print(
            f"📄 Лист: {sheet_name}"
        )

        try:

            ws = wb[sheet_name]

            rows = ws.iter_rows(
                values_only=True
            )

            article_col = None
            price_col = None

            header_row_number = None

            # Ищем заголовки в первых 30 строках
            for row_number in range(1, 31):

                try:
                    row = next(rows)

                except StopIteration:
                    break

                for index, value in enumerate(row):

                    header = normalize_header(value)

                    if header == normalize_header("Артикул"):
                        article_col = index

                    if header == normalize_header("Оптова ціна, $"):
                        price_col = index

                if (
                    article_col is not None
                    and price_col is not None
                ):

                    header_row_number = row_number
                    break

            if (
                article_col is None
                or price_col is None
            ):

                print(
                    "   ⚠️ Нужные колонки не найдены"
                )

                continue

            print(
                f"   ✅ Заголовки найдены "
                f"в строке {header_row_number}"
            )

            print(
                f"   🔹 Артикул: "
                f"колонка {article_col + 1}"
            )

            print(
                f"   🔹 Оптова ціна, $: "
                f"колонка {price_col + 1}"
            )

            sheet_count = 0

            # =========================
            # СТРОКИ
            # =========================

            for row in rows:

                if not row:
                    continue

                if article_col >= len(row):
                    continue

                if price_col >= len(row):
                    continue

                sku_raw = row[article_col]
                price_raw = row[price_col]

                sku = normalize_sku(
                    sku_raw
                )

                if not sku:
                    continue

                price = parse_price(
                    price_raw
                )

                if price is None:
                    continue

                sheet_count += 1

                if sku in prices:

                    duplicate_count += 1

                    print(
                        f"   ⚠️ Дубликат: "
                        f"{sku_raw} "
                        f"→ уже ${prices[sku]:g}, "
                        f"новая ${price:g}"
                    )

                    # Оставляем первую цену
                    continue

                prices[sku] = price

            print(
                f"   💰 Цен на листе: "
                f"{sheet_count}"
            )

        except Exception as e:

            print(
                f"   ❌ Ошибка листа: {e}"
            )

    print("")
    print("-" * 70)
    print(
        f"💰 Уникальных SKU с ценой: "
        f"{len(prices)}"
    )
    print(
        f"🔁 Дубликатов: "
        f"{duplicate_count}"
    )
    print("-" * 70)

    # =========================
    # ПОКАЗЫВАЕМ ПРИМЕРЫ
    # =========================

    if prices:

        print("")
        print("🔎 Примеры цен:")

        for i, (sku, price) in enumerate(
            prices.items()
        ):

            print(
                f"   {sku} → ${price:g}"
            )

            if i >= 19:
                break

    else:

        print("")
        print("❌ ЦЕНЫ НЕ НАЙДЕНЫ!")

    print("")

    return prices


# =========================
# CATEGORIES
# =========================

def get_categories():

    soup = get_soup(BASE)

    categories = []

    # Весь каталог находится здесь
    menu = soup.select_one(
        "#categories_block_left ul.tree"
    )

    if not menu:

        print(
            "⚠️ #categories_block_left "
            "не найден"
        )

        return categories

    seen = set()

    # Берём ВСЕ ссылки дерева,
    # независимо от уровня вложенности
    for a in menu.select("a[href]"):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:
            continue

        href = absolute_url(href)

        # Убираем hash
        href = href.split("#")[0]

        if href in seen:
            continue

        seen.add(href)

        name = clean(
            a.get_text(
                " ",
                strip=True
            )
        )

        if not name:
            continue

        categories.append({
            "name": name,
            "url": href
        })

    print(
        f"📂 Categories: "
        f"{len(categories)}"
    )

    for i, category in enumerate(
        categories,
        1
    ):

        print(
            f"   {i}. "
            f"{category['name']} "
            f"→ {category['url']}"
        )

    return categories


# =========================
# SHOW ALL
# =========================

def get_show_all_data(soup):

    form = soup.select_one(
        "form.showall"
    )

    if not form:
        return None

    n_input = form.select_one(
        "input[name='n']"
    )

    id_input = form.select_one(
        "input[name='id_category']"
    )

    if not n_input or not id_input:
        return None

    n = clean(
        n_input.get("value", "")
    )

    category_id = clean(
        id_input.get("value", "")
    )

    if not n or not category_id:
        return None

    return {
        "n": n,
        "id_category": category_id
    }


# =========================
# PRODUCT LINKS
# =========================

def get_product_links(soup):

    links = []

    seen = set()

    selectors = [

        "ul.product_list li.ajax_block_product a.product-name",

        ".product-container a.product-name",

        "li.product-box.item a.product-image",

        "li.product-box.item h5.product-name a",

        "li.product-box.item a[href*='.html']",

        "#productscategory_list li.product-box a[href*='.html']",

    ]

    for selector in selectors:

        for a in soup.select(selector):

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:
                continue

            href = absolute_url(
                href
            )

            href = clean_product_url(
                href
            )

            if not href:
                continue

            if href in seen:
                continue

            seen.add(href)

            links.append(href)

    return links


# =========================
# PAGINATION
# =========================

def get_last_page(soup):

    pages = [1]

    for a in soup.select(
        "#pagination_bottom ul.pagination a"
    ):

        href = a.get(
            "href",
            ""
        )

        # Например:
        # #/page-2
        m = re.search(
            r"#/page-(\d+)",
            href,
            re.I
        )

        if m:

            pages.append(
                int(m.group(1))
            )

    return max(pages)


# =========================
# PARSE CATEGORY
# =========================

def parse_category(cat_url):

    all_items = []

    seen_products = set()

    # =========================
    # ПЕРВАЯ СТРАНИЦА
    # =========================

    first_page = get_soup(
        cat_url
    )

    # =========================
    # SHOW ALL
    # =========================

    show_all = get_show_all_data(
        first_page
    )

    if show_all:

        show_all_url = (
            cat_url
            + "?n="
            + show_all["n"]
            + "&id_category="
            + show_all["id_category"]
        )

        print(
            f"📄 SHOW ALL: "
            f"{show_all_url}"
        )

        soup = get_soup(
            show_all_url
        )

        products = get_product_links(
            soup
        )

        for url in products:

            if url in seen_products:
                continue

            seen_products.add(url)
            all_items.append(url)

        if all_items:

            print(
                f"🛒 Products: "
                f"{len(all_items)}"
            )

            return all_items

    # =========================
    # ОБЫЧНАЯ ПАГИНАЦИЯ
    # =========================

    last_page = get_last_page(
        first_page
    )

    print(
        f"📄 Pages: "
        f"{last_page}"
    )

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:

            url = cat_url

        else:

            # ВАЖНО:
            # hash #/page-2 браузер обрабатывает
            # сам, requests его не отправляет.
            # Поэтому пробуем ?page=2
            separator = (
                "&"
                if "?" in cat_url
                else "?"
            )

            url = (
                cat_url
                + separator
                + f"page={page}"
            )

        soup = get_soup(
            url
        )

        products = get_product_links(
            soup
        )

        print(
            f"📄 Page {page}: "
            f"{len(products)} products"
        )

        for product_url in products:

            if product_url in seen_products:
                continue

            seen_products.add(
                product_url
            )

            all_items.append(
                product_url
            )

    return all_items


# =========================
# PRODUCT TITLE
# =========================

def get_product_title(soup):

    selectors = [

        "h1[itemprop='name']",

        "h1.product-name",

        "#product_name",

        "h1",

    ]

    for selector in selectors:

        element = soup.select_one(
            selector
        )

        if element:

            title = clean(
                element.get_text(
                    " ",
                    strip=True
                )
            )

            if title:
                return title

    return ""


# =========================
# PRODUCT SKU
# =========================

def get_product_sku(soup):

    selectors = [

        "#product_reference span[itemprop='sku']",

        "#product_reference .editable",

        "#product_reference",

        "[itemprop='sku']",

    ]

    for selector in selectors:

        element = soup.select_one(
            selector
        )

        if element:

            sku = clean(
                element.get_text(
                    " ",
                    strip=True
                )
            )

            if sku:

                sku = re.sub(
                    r"^(артикул|код|sku)"
                    r"\s*[:#-]?\s*",
                    "",
                    sku,
                    flags=re.I
                )

                return clean(sku)

    return ""


# =========================
# PRODUCT STATUS
# =========================

def get_product_status(soup):

    # =========================
    # Видимый статус
    # =========================

    element = soup.select_one(
        "#availability_value"
    )

    if element:

        text = clean(
            element.get_text(
                " ",
                strip=True
            )
        )

        low = text.casefold()

        unavailable = [

            "відсут",

            "немає",

            "нема",

            "нет в наличии",

            "відсутній",

            "відсутня",

        ]

        for word in unavailable:

            if word in low:
                return "Немає в наявності"

        return "В наявності"

    # =========================
    # Schema
    # =========================

    availability = soup.select_one(
        "link[itemprop='availability']"
    )

    if availability:

        value = (
            availability.get("href")
            or availability.get("content")
            or ""
        )

        value = str(
            value
        ).casefold()

        if "outofstock" in value:
            return "Немає в наявності"

        if "instock" in value:
            return "В наявності"

    return "В наявності"


# =========================
# COLORS
# =========================

def get_product_colors(soup):

    colors = []

    # Ищем только fieldset,
    # где legend содержит "Колір"
    fieldsets = soup.select(
        "#attributes .attribute_fieldset"
    )

    for fieldset in fieldsets:

        legend = fieldset.select_one(
            "legend"
        )

        if not legend:
            continue

        legend_text = clean(
            legend.get_text(
                " ",
                strip=True
            )
        ).casefold()

        if "колір" not in legend_text:
            continue

        inputs = fieldset.select(
            "input.attribute_radio"
        )

        for input_el in inputs:

            color_id = input_el.get(
                "value",
                ""
            ).strip()

            if not color_id:
                continue

            color_name = ""

            # =========================
            # LABEL
            # =========================

            label = None

            input_id = input_el.get(
                "id"
            )

            if input_id:

                label = fieldset.select_one(
                    f"label[for='{input_id}']"
                )

            if label:

                color_name = clean(
                    label.get_text(
                        " ",
                        strip=True
                    )
                )

            # =========================
            # PARENT
            # =========================

            if not color_name:

                parent = input_el.parent

                if parent:

                    color_name = clean(
                        parent.get_text(
                            " ",
                            strip=True
                        )
                    )

            # =========================
            # LI
            # =========================

            if not color_name:

                li = input_el.find_parent(
                    "li"
                )

                if li:

                    color_name = clean(
                        li.get_text(
                            " ",
                            strip=True
                        )
                    )

            color_name = re.sub(
                r"\bchecked\b",
                "",
                color_name,
                flags=re.I
            )

            color_name = clean(
                color_name
            )

            if not color_name:
                continue

            colors.append({
                "id": color_id,
                "name": color_name
            })

    # =========================
    # DEDUPE COLORS
    # =========================

    result = []

    seen = set()

    for color in colors:

        key = (
            color["id"],
            normalize_sku(
                color["name"]
            )
        )

        if key in seen:
            continue

        seen.add(key)

        result.append(
            color
        )

    return result


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

    if not soup:
        return []

    title = get_product_title(
        soup
    )

    sku = get_product_sku(
        soup
    )

    status = get_product_status(
        soup
    )

    # =========================
    # PRICE FROM GOOGLE SHEETS
    # =========================

    normalized_sku = normalize_sku(
        sku
    )

    price = price_map.get(
        normalized_sku
    )

    if price is None:

        price = ""

    # =========================
    # COLORS
    # =========================

    colors = get_product_colors(
        soup
    )

    result = []

    # =========================
    # БЕЗ ЦВЕТОВ
    # =========================

    if not colors:

        result.append([
            sku,
            title,
            price,
            status,
            "",
            product_url
        ])

        return result

    # =========================
    # КАЖДЫЙ ЦВЕТ ОТДЕЛЬНО
    # =========================

    for color in colors:

        color_id = color["id"]
        color_name = color["name"]

        # Ссылка именно на выбранный цвет
        color_url = (
            product_url
            + "#/"
            + str(color_id)
            + "-kolir-"
            + color_name
        )

        result.append([
            sku,
            title,
            price,
            status,
            color_name,
            color_url
        ])

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
        # GOOGLE SHEETS
        # =========================

        price_map = get_google_sheet_price_map()

        if not price_map:

            print(
                "❌ Цены из Google Sheets "
                "не загружены."
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # =========================
        # EXCEL
        # =========================

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE USD",
            "STATUS",
            "COLOR",
            "URL"
        ])

        seen = set()

        # =========================
        # CATEGORIES
        # =========================

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

        # =========================
        # CATEGORIES
        # =========================

        for i, cat in enumerate(
            cats,
            1
        ):

            save_status(
                True,
                int(i / total * 100),
                USER,
                FILE_PATH
            )

            print("")
            print(
                f"📂 {i}/{total} "
                f"{cat['name']}"
            )

            product_urls = parse_category(
                cat["url"]
            )

            # =========================
            # PRODUCT LIMIT
            # =========================

            if PRODUCT_LIMIT:

                product_urls = product_urls[
                    :PRODUCT_LIMIT
                ]

            print(
                f"🛒 Товаров: "
                f"{len(product_urls)}"
            )

            # =========================
            # PRODUCTS
            # =========================

            for product_url in product_urls:

                # =========================
                # Защита от повторов
                # =========================

                if product_url in seen:
                    continue

                seen.add(
                    product_url
                )

                rows = parse_product(
                    product_url,
                    price_map
                )

                for row in rows:

                    sku = row[0]
                    title = row[1]
                    price = row[2]
                    status = row[3]
                    color = row[4]
                    url = row[5]

                    if not title:
                        continue

                    ws.append([
                        sku,
                        title,
                        price,
                        status,
                        color,
                        url
                    ])

                    print(
                        f"   ✓ {sku} | "
                        f"{color} | "
                        f"${price}"
                    )

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
            "✅ Готово. "
            "Самокаты ARIZONE Sports"
        )

    finally:

        set_lock(False)


# =========================
# START
# =========================

if __name__ == "__main__":
    run_parser()
