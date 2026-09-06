import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
import sys
from urllib.parse import urljoin, urlparse, urlunparse


# ============================================================
# НАСТРОЙКИ
# ============================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "Manual"

BASE = "https://arizonesports.com.ua/uk"
DOMAIN = "arizonesports.com.ua"

# Для первого теста можно поставить:
CATEGORY_LIMIT = 2
PRODUCT_LIMIT = 20
#
# После проверки:
#CATEGORY_LIMIT = None
#PRODUCT_LIMIT = None


OUTPUT_DIR = os.path.abspath("output/ArizoneSports")
FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "ArizoneSports_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "parser.lock"
)


# ============================================================
# HTTP
# ============================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7",
    "Connection": "keep-alive",
}


session = requests.Session()
session.headers.update(HEADERS)


# ============================================================
# LOCK
# ============================================================

def create_lock():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(LOCK_FILE):
        try:
            mtime = os.path.getmtime(LOCK_FILE)
            age = time.time() - mtime

            # Если lock старше часа — считаем его зависшим
            if age > 3600:
                os.remove(LOCK_FILE)
            else:
                print("⛔ PARSER ALREADY RUNNING")
                return False

        except Exception:
            pass

    try:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))

        return True

    except Exception as e:
        print(f"❌ LOCK ERROR: {e}")
        return False


def remove_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


# ============================================================
# STATUS
# ============================================================

def save_status(
    status,
    current=0,
    total=0,
    message=""
):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "status": status,
        "current": current,
        "total": total,
        "message": message,
        "updated": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    try:
        with open(
            STATUS_PATH,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )
    except Exception:
        pass


# ============================================================
# HTTP GET
# ============================================================

def get_soup(url, params=None):
    try:
        response = session.get(
            url,
            params=params,
            timeout=30
        )

        response.raise_for_status()

        response.encoding = response.apparent_encoding

        return BeautifulSoup(
            response.text,
            "html.parser"
        )

    except Exception as e:
        print(f"❌ GET ERROR: {url}")
        print(f"   {e}")
        return None


# ============================================================
# URL
# ============================================================

def absolute_url(url):
    if not url:
        return ""

    url = url.strip()

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("http://") or url.startswith("https://"):
        return url

    return urljoin(
        BASE + "/",
        url
    )


def base_product_url(url):
    """
    Убирает #цвет из URL.

    Например:

    /187-product.html#/43-kolir-bezhevij

    превращается в:

    /187-product.html
    """

    parsed = urlparse(url)

    return urlunparse(
        (
            parsed.scheme,
            parsed.netloc,
            parsed.path,
            parsed.params,
            parsed.query,
            ""
        )
    )


# ============================================================
# CLEAN
# ============================================================

def clean(value):
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value.strip()


# ============================================================
# ТРАНСЛИТЕРАЦИЯ ЦВЕТОВ
# ============================================================

TRANSLIT = {
    "а": "a",
    "б": "b",
    "в": "v",
    "г": "g",
    "ґ": "g",
    "д": "d",
    "е": "e",
    "є": "ye",
    "ж": "zh",
    "з": "z",
    "и": "i",
    "і": "i",
    "ї": "yi",
    "й": "j",
    "к": "k",
    "л": "l",
    "м": "m",
    "н": "n",
    "о": "o",
    "п": "p",
    "р": "r",
    "с": "s",
    "т": "t",
    "у": "u",
    "ф": "f",
    "х": "kh",
    "ц": "ts",
    "ч": "ch",
    "ш": "sh",
    "щ": "shch",
    "ь": "",
    "ю": "yu",
    "я": "ya",

    "ё": "yo",
    "ы": "y",
    "э": "e",
    "ъ": "",

    "А": "A",
    "Б": "B",
    "В": "V",
    "Г": "G",
    "Ґ": "G",
    "Д": "D",
    "Е": "E",
    "Є": "Ye",
    "Ж": "Zh",
    "З": "Z",
    "И": "I",
    "І": "I",
    "Ї": "Yi",
    "Й": "J",
    "К": "K",
    "Л": "L",
    "М": "M",
    "Н": "N",
    "О": "O",
    "П": "P",
    "Р": "R",
    "С": "S",
    "Т": "T",
    "У": "U",
    "Ф": "F",
    "Х": "Kh",
    "Ц": "Ts",
    "Ч": "Ch",
    "Ш": "Sh",
    "Щ": "Shch",
    "Ь": "",
    "Ю": "Yu",
    "Я": "Ya",

    "Ё": "Yo",
    "Ы": "Y",
    "Э": "E",
    "Ъ": "",
}


def translit_slug(text):
    """
    Делает slug примерно в том же формате,
    который использует сайт в hash URL.

    Бежевий -> bezhevij
    Білий   -> bilij
    Зелений -> zelenij
    Рожевий -> rozhevij
    Чорний  -> chornij
    """

    text = clean(text)

    result = []

    for char in text:
        result.append(
            TRANSLIT.get(char, char)
        )

    text = "".join(result)

    text = text.lower()

    # Сайт использует дефисы
    text = re.sub(
        r"[^a-z0-9]+",
        "-",
        text
    )

    text = text.strip("-")

    return text


# ============================================================
# КАТЕГОРИИ
# ============================================================

def get_categories():
    print("📂 GET CATEGORIES")

    soup = get_soup(BASE)

    if not soup:
        return []

    container = soup.select_one(
        "#categories_block_left"
    )

    if not container:
        print("❌ #categories_block_left NOT FOUND")
        return []

    categories = []

    # ВАЖНО:
    #
    # Не ограничиваемся первым уровнем.
    #
    # Берём ВСЕ ссылки внутри дерева:
    #
    # категория
    # └── подкатегория
    #     └── подподкатегория
    #
    links = container.select(
        "ul.tree a[href]"
    )

    seen = set()

    for a in links:
        href = a.get("href")

        if not href:
            continue

        url = absolute_url(href)

        parsed = urlparse(url)

        if parsed.netloc and parsed.netloc != DOMAIN:
            continue

        # Убираем hash/query
        url = base_product_url(url)

        if url in seen:
            continue

        seen.add(url)

        name = clean(
            a.get_text(
                " ",
                strip=True
            )
        )

        if not name:
            continue

        categories.append(
            {
                "name": name,
                "url": url
            }
        )

    print(
        f"📂 FOUND CATEGORIES: {len(categories)}"
    )

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

    return categories


# ============================================================
# SHOW ALL
# ============================================================

def get_show_all_data(cat_url, soup):
    """
    На сайте есть:

    <form class="showall" ...>
        <input name="n" value="85">
        <input name="id_category" value="14">
    </form>

    Используем этот механизм, чтобы получить
    сразу все товары категории.
    """

    form = soup.select_one(
        "form.showall"
    )

    if not form:
        return None

    action = form.get("action")

    if not action:
        action = cat_url

    action = absolute_url(action)

    params = {}

    for inp in form.select(
        "input[name]"
    ):
        name = inp.get("name")
        value = inp.get("value", "")

        if name:
            params[name] = value

    if "n" not in params:
        return None

    try:
        count = int(
            re.sub(
                r"\D",
                "",
                str(params["n"])
            )
        )
    except Exception:
        count = 0

    if count <= 0:
        return None

    return {
        "url": action,
        "params": params,
        "count": count
    }


# ============================================================
# PRODUCT LINKS
# ============================================================

def extract_product_links(soup):
    """
    Достаём ссылки именно на товары.

    Основной шаблон сайта:

    <li class="product-box item">
        <a class="lnk_img product-image"
           href="...html">
    """

    if not soup:
        return []

    links = []

    seen = set()

    selectors = [
        "ul.product_list li.ajax_block_product a.product-name",
        "ul.product_list li.ajax_block_product h5 a",
        "ul.product_list li.ajax_block_product a[href*='.html']",

        ".product-container a.product-name",
        ".product-container h5 a",
        ".product-container a[href*='.html']",

        "li.product-box.item a.product-image",
        "li.product-box.item h5.product-name a",
        "li.product-box.item a[href*='.html']",

        "#productscategory_list li.product-box a[href*='.html']",
    ]

    for selector in selectors:

        for a in soup.select(selector):

            href = a.get("href")

            if not href:
                continue

            href = absolute_url(href)

            parsed = urlparse(href)

            if parsed.netloc != DOMAIN:
                continue

            path = parsed.path.lower()

            if not path.endswith(".html"):
                continue

            # Убираем # и query
            href = base_product_url(href)

            if href in seen:
                continue

            seen.add(href)

            links.append(href)

    return links


# ============================================================
# CATEGORY PARSER
# ============================================================

def parse_category(category):
    name = category["name"]
    cat_url = category["url"]

    print()
    print(
        f"📁 CATEGORY: {name}"
    )
    print(
        f"🔗 {cat_url}"
    )

    soup = get_soup(cat_url)

    if not soup:
        return []

    # --------------------------------------------------------
    # Сначала пробуем "Показати все"
    # --------------------------------------------------------

    show_all = get_show_all_data(
        cat_url,
        soup
    )

    if show_all:

        print(
            f"📦 CATEGORY PRODUCTS EXPECTED: "
            f"{show_all['count']}"
        )

        all_soup = get_soup(
            show_all["url"],
            params=show_all["params"]
        )

        if all_soup:

            links = extract_product_links(
                all_soup
            )

            print(
                f"🔎 SHOW ALL PRODUCT LINKS: "
                f"{len(links)}"
            )

            # Если получили товары — используем
            if links:

                # Иногда сервер может вернуть не все товары.
                # Поэтому, если количество выглядит разумно,
                # принимаем результат.
                if (
                    show_all["count"] <= 0
                    or len(links) >= show_all["count"]
                    or len(links) > 1
                ):
                    return links

    # --------------------------------------------------------
    # Если SHOW ALL не сработал
    # --------------------------------------------------------
    #
    # У сайта пагинация:
    #
    # /category.html#/page-2
    #
    # Но # сервер не получает.
    #
    # Поэтому сначала пробуем обычные ссылки ?page=N
    # и дополнительно ищем количество страниц.
    # --------------------------------------------------------

    first_links = extract_product_links(
        soup
    )

    print(
        f"🔎 FIRST PAGE PRODUCT LINKS: "
        f"{len(first_links)}"
    )

    links = list(first_links)

    seen = set(links)

    # Ищем все номера страниц
    page_numbers = set()

    for a in soup.select(
        "#pagination_bottom a[href]"
    ):
        href = a.get("href", "")

        match = re.search(
            r"#/page-(\d+)",
            href
        )

        if match:
            page_numbers.add(
                int(match.group(1))
            )

    if page_numbers:

        max_page = max(
            page_numbers
        )

        print(
            f"📄 PAGES FOUND: 1-{max_page}"
        )

        for page in range(
            2,
            max_page + 1
        ):

            # PrestaShop часто принимает page
            # через query parameter.
            params = {
                "page": page
            }

            page_soup = get_soup(
                cat_url,
                params=params
            )

            if not page_soup:
                continue

            page_links = extract_product_links(
                page_soup
            )

            print(
                f"   PAGE {page}: "
                f"{len(page_links)} products"
            )

            for link in page_links:

                if link not in seen:
                    seen.add(link)
                    links.append(link)

    print(
        f"📦 TOTAL CATEGORY PRODUCTS: "
        f"{len(links)}"
    )

    return links


# ============================================================
# STATUS
# ============================================================

def parse_status(soup):
    """
    Варианты на сайте:

    1.
    <span id="availability_value"
          class="label label-danger">
        Наразі товар відсутній
    </span>

    2.
    availability_value пустой,
    но есть:

    <link itemprop="availability"
          href="https://schema.org/InStock">
    """

    if not soup:
        return ""

    # --------------------------------------------------------
    # 1. Видимый текст
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # 2. Schema.org
    # --------------------------------------------------------

    schema = soup.select_one(
        "link[itemprop='availability']"
    )

    if schema:

        href = (
            schema.get("href", "")
            .strip()
        )

        if "InStock" in href:
            return "В наявності"

        if "OutOfStock" in href:
            return "Наразі товар відсутній"

        if "PreOrder" in href:
            return "Під замовлення"

    return ""


# ============================================================
# SKU
# ============================================================

def parse_sku(soup):
    """
    Артикул:

    <p id="product_reference">
        <label>Артикул</label>
        <span itemprop="sku"
              content="BBTT-Q7">
            BBTT-Q7
        </span>
    </p>
    """

    if not soup:
        return ""

    sku = soup.select_one(
        "#product_reference span[itemprop='sku']"
    )

    if sku:

        value = (
            sku.get("content")
            or sku.get_text(
                " ",
                strip=True
            )
        )

        return clean(value)

    # Запасной вариант
    sku = soup.select_one(
        "#product_reference .editable"
    )

    if sku:
        return clean(
            sku.get_text(
                " ",
                strip=True
            )
        )

    return ""


# ============================================================
# TITLE
# ============================================================

def parse_title(soup):
    if not soup:
        return ""

    title = soup.select_one(
        "h1[itemprop='name']"
    )

    if title:
        return clean(
            title.get_text(
                " ",
                strip=True
            )
        )

    return ""


# ============================================================
# COLORS
# ============================================================

def parse_colors(soup):
    """
    На сайте:

    <fieldset class="attribute_fieldset">
        <label class="attribute_label">
            Колір
        </label>

        <div class="attribute_list">
            <ul>

                <li>
                    ...
                    <input
                        name="group_4"
                        value="43">
                    ...
                    <span>Бежевий</span>
                </li>

                ...

            </ul>
        </div>
    </fieldset>


    Возвращаем:

    [
        {
            "id": "43",
            "name": "Бежевий",
            "slug": "bezhevij"
        },
        ...
    ]
    """

    if not soup:
        return []

    colors = []

    seen = set()

    fieldsets = soup.select(
        "#attributes .attribute_fieldset"
    )

    for fieldset in fieldsets:

        label = fieldset.select_one(
            ".attribute_label"
        )

        if not label:
            continue

        label_text = clean(
            label.get_text(
                " ",
                strip=True
            )
        ).lower()

        # Ищем именно атрибут "Колір"
        if "колір" not in label_text:
            continue

        for li in fieldset.select(
            ".attribute_list li"
        ):

            input_el = li.select_one(
                "input.attribute_radio"
            )

            if not input_el:
                continue

            color_id = clean(
                input_el.get("value")
            )

            # Берём название цвета
            # именно из span, чтобы не захватывать
            # лишний текст.
            color_span = li.select_one(
                "span:not(.checked)"
            )

            if color_span:
                color_name = clean(
                    color_span.get_text(
                        " ",
                        strip=True
                    )
                )
            else:
                # Запасной вариант
                spans = li.select("span")

                color_name = ""

                for span in spans:
                    text = clean(
                        span.get_text(
                            " ",
                            strip=True
                        )
                    )

                    if text and text.lower() not in (
                        "checked",
                    ):
                        color_name = text

            if not color_name:
                continue

            key = (
                color_id,
                color_name.lower()
            )

            if key in seen:
                continue

            seen.add(key)

            colors.append(
                {
                    "id": color_id,
                    "name": color_name,
                    "slug": translit_slug(
                        color_name
                    )
                }
            )

    return colors


# ============================================================
# COLOR URL
# ============================================================

def make_color_url(
    product_url,
    color
):
    """
    Формируем URL именно так,
    как его делает сайт:

    ...html#/43-kolir-bezhevij

    ...html#/31-kolir-bilij

    ...html#/26-kolir-zelenij
    """

    product_url = base_product_url(
        product_url
    )

    color_id = color["id"]
    color_slug = color["slug"]

    return (
        f"{product_url}"
        f"#/{color_id}-kolir-{color_slug}"
    )


# ============================================================
# PRODUCT
# ============================================================

def parse_product(product_url):
    print(
        f"   🔎 PRODUCT: {product_url}"
    )

    soup = get_soup(
        product_url
    )

    if not soup:
        return []

    title = parse_title(
        soup
    )

    sku = parse_sku(
        soup
    )

    status = parse_status(
        soup
    )

    colors = parse_colors(
        soup
    )

    # --------------------------------------------------------
    # Если цветов НЕТ
    # создаём одну строку без цвета.
    # --------------------------------------------------------

    if not colors:

        return [
            {
                "SKU": sku,
                "TITLE": title,
                "PRICE USD": "",
                "STATUS": status,
                "COLOR": "",
                "URL": base_product_url(
                    product_url
                )
            }
        ]

    # --------------------------------------------------------
    # Если есть цвета:
    #
    # КАЖДЫЙ ЦВЕТ = ОТДЕЛЬНАЯ СТРОКА
    # --------------------------------------------------------

    products = []

    for color in colors:

        color_url = make_color_url(
            product_url,
            color
        )

        products.append(
            {
                "SKU": sku,
                "TITLE": title,
                "PRICE USD": "",
                "STATUS": status,
                "COLOR": color["name"],
                "URL": color_url
            }
        )

    return products


# ============================================================
# MAIN PARSER
# ============================================================

def run_parser():

    print("🔥 ARIZONESPORTS PARSER STARTED")
    print(f"👤 USER: {USER}")

    if not create_lock():
        return

    try:

        save_status(
            "starting",
            0,
            0,
            "Запуск парсера"
        )

        # ----------------------------------------------------
        # CATEGORIES
        # ----------------------------------------------------

        categories = get_categories()

        if not categories:

            print(
                "❌ CATEGORIES NOT FOUND"
            )

            save_status(
                "error",
                0,
                0,
                "Категории не найдены"
            )

            return

        # ----------------------------------------------------
        # COLLECT PRODUCT URLS
        # ----------------------------------------------------

        product_urls = []

        seen_urls = set()

        total_categories = len(
            categories
        )

        for category_index, category in enumerate(
            categories,
            start=1
        ):

            save_status(
                "categories",
                category_index,
                total_categories,
                category["name"]
            )

            print()
            print(
                f"📂 [{category_index}/"
                f"{total_categories}] "
                f"{category['name']}"
            )

            links = parse_category(
                category
            )

            for url in links:

                url = base_product_url(
                    url
                )

                if url in seen_urls:
                    continue

                seen_urls.add(url)

                product_urls.append(
                    url
                )

                if (
                    PRODUCT_LIMIT
                    and
                    len(product_urls)
                    >= PRODUCT_LIMIT
                ):
                    break

            if (
                PRODUCT_LIMIT
                and
                len(product_urls)
                >= PRODUCT_LIMIT
            ):
                break

        print()
        print(
            f"🛒 UNIQUE PRODUCTS: "
            f"{len(product_urls)}"
        )

        if not product_urls:

            print(
                "❌ PRODUCTS NOT FOUND"
            )

            save_status(
                "error",
                0,
                0,
                "Товары не найдены"
            )

            return

        # ----------------------------------------------------
        # EXCEL
        # ----------------------------------------------------

        wb = Workbook()

        ws = wb.active
        ws.title = "Products"

        headers = [
            "SKU",
            "TITLE",
            "PRICE USD",
            "STATUS",
            "COLOR",
            "URL"
        ]

        ws.append(
            headers
        )

        # ----------------------------------------------------
        # PARSE PRODUCTS
        # ----------------------------------------------------

        total_products = len(
            product_urls
        )

        # Теперь дедупликация:
        #
        # SKU + COLOR
        #
        # а НЕ SKU отдельно.
        #
        # Потому что:
        #
        # BBTT-Q7 + Бежевий
        # BBTT-Q7 + Білий
        # BBTT-Q7 + Зелений
        #
        # это три строки.
        seen_variants = set()

        parsed_count = 0

        for index, product_url in enumerate(
            product_urls,
            start=1
        ):

            save_status(
                "products",
                index,
                total_products,
                product_url
            )

            rows = parse_product(
                product_url
            )

            for row in rows:

                sku = clean(
                    row.get("SKU")
                )

                color = clean(
                    row.get("COLOR")
                )

                # ------------------------------------------------
                # Дедупликация
                # ------------------------------------------------
                #
                # Если SKU есть:
                # SKU + COLOR
                #
                # Если SKU почему-то нет:
                # URL + COLOR
                #
                # ------------------------------------------------

                if sku:

                    variant_key = (
                        sku.lower(),
                        color.lower()
                    )

                else:

                    variant_key = (
                        row["URL"].split("#")[0].lower(),
                        color.lower()
                    )

                if variant_key in seen_variants:
                    continue

                seen_variants.add(
                    variant_key
                )

                ws.append(
                    [
                        row["SKU"],
                        row["TITLE"],
                        row["PRICE USD"],
                        row["STATUS"],
                        row["COLOR"],
                        row["URL"]
                    ]
                )

                parsed_count += 1

            print(
                f"   ✅ [{index}/"
                f"{total_products}] "
                f"rows: {len(rows)}"
            )

        # ----------------------------------------------------
        # WIDTH
        # ----------------------------------------------------

        widths = {
            "A": 18,
            "B": 65,
            "C": 15,
            "D": 30,
            "E": 20,
            "F": 100
        }

        for column, width in widths.items():
            ws.column_dimensions[
                column
            ].width = width

        # ----------------------------------------------------
        # FREEZE
        # ----------------------------------------------------

        ws.freeze_panes = "A2"

        # ----------------------------------------------------
        # SAVE
        # ----------------------------------------------------

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        temp_file = FILE_PATH + ".tmp"

        wb.save(
            temp_file
        )

        # Если старый файл существует —
        # заменяем его новым.
        if os.path.exists(
            FILE_PATH
        ):
            os.remove(
                FILE_PATH
            )

        os.replace(
            temp_file,
            FILE_PATH
        )

        print()
        print(
            "========================================"
        )
        print(
            "🔥 ARIZONESPORTS PARSER FINISHED"
        )
        print(
            f"📂 Categories: "
            f"{len(categories)}"
        )
        print(
            f"🛒 Product URLs: "
            f"{len(product_urls)}"
        )
        print(
            f"📄 Excel rows: "
            f"{parsed_count}"
        )
        print(
            f"💾 FILE: "
            f"{FILE_PATH}"
        )
        print(
            "========================================"
        )

        save_status(
            "finished",
            parsed_count,
            parsed_count,
            "Готово"
        )

    except Exception as e:

        print()
        print(
            "❌ PARSER ERROR"
        )

        print(
            repr(e)
        )

        save_status(
            "error",
            0,
            0,
            str(e)
        )

        raise

    finally:

        remove_lock()


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    run_parser()
