import os
import json
import re
import time
import csv
from io import StringIO
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
import sys


# ============================================================
# НАСТРОЙКИ
# ============================================================

BASE = "https://arizonesports.com.ua/uk"
DOMAIN = "arizonesports.com.ua"

# Google Таблица
GOOGLE_SHEET_ID = "1ROCGu-3W8PotpQSIgzHRDkXCiDeHU0_dzEjuRnkNHv0"

# Для теста можно поставить число.
# None = обрабатывать всё
CATEGORY_LIMIT = 2
PRODUCT_LIMIT = 20

OUTPUT_DIR = os.path.abspath("output/ARIZONE_Sports")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "ARIZONE_Sports_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "parser.lock"
)

USER = sys.argv[1] if len(sys.argv) > 1 else "unknown"


# ============================================================
# SESSION
# ============================================================

session = requests.Session()

session.headers.update({
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7",
})


# ============================================================
# LOCK
# ============================================================

def create_lock():

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(LOCK_FILE):

        try:
            age = time.time() - os.path.getmtime(LOCK_FILE)

            # Старый lock считаем зависшим
            if age > 3600:
                os.remove(LOCK_FILE)
                print("⚠️ Удалён старый lock")

            else:
                print("❌ Парсер уже запущен")
                return False

        except Exception:
            pass

    try:

        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))

        return True

    except Exception as e:

        print(f"❌ Не удалось создать lock: {e}")
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
        "updated": datetime.now().isoformat()
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

    except Exception as e:

        print(f"⚠️ Ошибка status.json: {e}")


# ============================================================
# HTTP
# ============================================================

def get_soup(url, params=None, retries=3):

    for attempt in range(1, retries + 1):

        try:

            response = session.get(
                url,
                params=params,
                timeout=30
            )

            response.raise_for_status()

            return BeautifulSoup(
                response.text,
                "html.parser"
            )

        except Exception as e:

            print(
                f"⚠️ Ошибка GET "
                f"{url} "
                f"(попытка {attempt}/{retries}): {e}"
            )

            if attempt < retries:
                time.sleep(2)

    return None


# ============================================================
# HELPERS
# ============================================================

def absolute_url(url):

    if not url:
        return ""

    url = url.strip()

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("/"):
        return "https://" + DOMAIN + url

    if url.startswith("http://"):
        return url.replace(
            "http://",
            "https://",
            1
        )

    if url.startswith("https://"):
        return url

    return BASE.rstrip("/") + "/" + url.lstrip("/")


def base_product_url(url):

    if not url:
        return ""

    url = url.split("#")[0]
    url = url.split("?")[0]

    return url


def clean(value):

    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
    ).strip()


# ============================================================
# TRANSLITERATION
# ============================================================

TRANSLIT_MAP = {
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
    "ы": "y",
    "э": "e",
    "ъ": "",
}


def translit_slug(value):

    value = clean(value).lower()

    result = ""

    for char in value:

        if char in TRANSLIT_MAP:
            result += TRANSLIT_MAP[char]

        elif (
            char.isalnum()
            or char in "-_"
        ):
            result += char

        elif char in " .":
            result += "-"

    result = re.sub(
        r"-+",
        "-",
        result
    )

    return result.strip("-")


# ============================================================
# GOOGLE SHEETS
# ============================================================

def normalize_header(value):

    if value is None:
        return ""

    value = str(value).strip().lower()

    value = value.replace(
        "\n",
        " "
    )

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    return value


def normalize_sku(value):

    if value is None:
        return ""

    value = str(value).strip()

    # Например Google может вернуть 1234.0
    if re.fullmatch(
        r"\d+\.0",
        value
    ):
        value = value[:-2]

    return value.strip().lower()


def parse_price(value):

    if value is None:
        return None

    value = str(value).strip()

    if not value:
        return None

    value = value.replace(
        "$",
        ""
    )

    value = value.replace(
        "USD",
        ""
    )

    value = value.replace(
        "usd",
        ""
    )

    value = value.strip()

    # 28,50 → 28.50
    if "," in value and "." not in value:
        value = value.replace(
            ",",
            "."
        )

    # 1 250 → 1250
    value = value.replace(
        " ",
        ""
    )

    try:

        return float(value)

    except ValueError:

        return None


def get_google_sheet_price_map():

    print()
    print("=" * 60)
    print("📊 ЗАГРУЗКА ЦЕН ИЗ GOOGLE SHEETS")
    print("=" * 60)

    # --------------------------------------------------------
    # Получаем HTML Google Таблицы
    # --------------------------------------------------------

    metadata_url = (
        "https://docs.google.com/spreadsheets/d/"
        f"{GOOGLE_SHEET_ID}"
        "/gviz/tq?tqx=out:html"
    )

    try:

        response = session.get(
            metadata_url,
            timeout=30
        )

        response.raise_for_status()

    except Exception as e:

        print(
            f"❌ Не удалось открыть Google Таблицу: {e}"
        )

        return {}

    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )

    # --------------------------------------------------------
    # Ищем все gid вкладок
    # --------------------------------------------------------

    sheet_links = []

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = a.get(
            "href",
            ""
        )

        if "gid=" not in href:
            continue

        match = re.search(
            r"gid=(\d+)",
            href
        )

        if not match:
            continue

        gid = match.group(1)

        name = clean(
            a.get_text(
                " ",
                strip=True
            )
        )

        if not name:
            name = f"Лист {gid}"

        exists = False

        for old_gid, old_name in sheet_links:

            if old_gid == gid:
                exists = True
                break

        if not exists:

            sheet_links.append(
                (
                    gid,
                    name
                )
            )

    if not sheet_links:

        print(
            "❌ Не удалось получить список листов Google Таблицы."
        )

        print(
            "Проверь, что таблица открыта "
            "для просмотра по ссылке."
        )

        return {}

    print(
        f"📚 Найдено листов: {len(sheet_links)}"
    )

    # --------------------------------------------------------
    # Общий словарь цен
    # --------------------------------------------------------

    prices = {}

    # Дубликаты
    duplicates = []

    # --------------------------------------------------------
    # Идём по каждому листу
    # --------------------------------------------------------

    for gid, sheet_name in sheet_links:

        print()
        print(
            f"📄 Лист: {sheet_name} "
            f"(gid={gid})"
        )

        csv_url = (
            "https://docs.google.com/spreadsheets/d/"
            f"{GOOGLE_SHEET_ID}"
            f"/export?format=csv&gid={gid}"
        )

        try:

            r = session.get(
                csv_url,
                timeout=30
            )

            r.raise_for_status()

            text = r.content.decode(
                "utf-8-sig"
            )

        except Exception as e:

            print(
                f"   ❌ Ошибка загрузки листа: {e}"
            )

            continue

        # ----------------------------------------------------
        # CSV
        # ----------------------------------------------------

        reader = csv.reader(
            StringIO(text)
        )

        rows = list(reader)

        if not rows:

            print(
                "   ⚠️ Лист пустой"
            )

            continue

        # ----------------------------------------------------
        # Ищем строку заголовков
        # ----------------------------------------------------

        header_index = None
        article_col = None
        price_col = None

        for i, row in enumerate(
            rows[:15]
        ):

            normalized_headers = [
                normalize_header(x)
                for x in row
            ]

            found_article = None
            found_price = None

            for col, header in enumerate(
                normalized_headers
            ):

                if header == "артикул":
                    found_article = col

                if header == "оптова ціна, $":
                    found_price = col

            if (
                found_article is not None
                and
                found_price is not None
            ):

                header_index = i
                article_col = found_article
                price_col = found_price

                break

        # ----------------------------------------------------
        # Если колонки не найдены
        # ----------------------------------------------------

        if header_index is None:

            print(
                "   ⚠️ Не найдены колонки:"
            )

            print(
                "      Артикул"
            )

            print(
                "      Оптова ціна, $"
            )

            continue

        print(
            f"   ✅ Артикул: колонка "
            f"{article_col + 1}"
        )

        print(
            f"   ✅ Оптова ціна, $: колонка "
            f"{price_col + 1}"
        )

        # ----------------------------------------------------
        # Читаем строки
        # ----------------------------------------------------

        added = 0

        for row in rows[
            header_index + 1:
        ]:

            if article_col >= len(row):
                continue

            if price_col >= len(row):
                continue

            sku_raw = row[
                article_col
            ]

            price_raw = row[
                price_col
            ]

            sku = normalize_sku(
                sku_raw
            )

            price = parse_price(
                price_raw
            )

            if not sku:
                continue

            if price is None:
                continue

            # ------------------------------------------------
            # SKU уже есть
            # ------------------------------------------------

            if sku in prices:

                old_price = prices[
                    sku
                ]

                if old_price != price:

                    duplicates.append({
                        "sku": sku,
                        "old_price": old_price,
                        "new_price": price,
                        "sheet": sheet_name
                    })

                    print(
                        f"   ⚠️ Дубликат "
                        f"{sku}: "
                        f"{old_price} / {price}"
                    )

                    # Оставляем первое найденное
                    continue

            else:

                prices[sku] = price
                added += 1

        print(
            f"   ➕ Добавлено цен: {added}"
        )

    # --------------------------------------------------------
    # Результат
    # --------------------------------------------------------

    print()
    print(
        f"💰 Всего уникальных SKU с ценой: "
        f"{len(prices)}"
    )

    if duplicates:

        print()
        print(
            "⚠️ ОБНАРУЖЕНЫ ДУБЛИКАТЫ "
            "SKU С РАЗНЫМИ ЦЕНАМИ:"
        )

        for item in duplicates:

            print(
                f"   {item['sku']}: "
                f"{item['old_price']} / "
                f"{item['new_price']} "
                f"(лист: {item['sheet']})"
            )

    print()
    print(
        "✅ Google Sheets загружена"
    )

    print(
        "=" * 60
    )
    print()

    return prices


# ============================================================
# КАТЕГОРИИ
# ============================================================

def get_categories():

    print("📂 Получаем категории...")

    soup = get_soup(BASE)

    if not soup:
        return []

    block = soup.select_one(
        "#categories_block_left"
    )

    if not block:

        print(
            "❌ Блок категорий не найден"
        )

        return []

    categories = []
    seen = set()

    # --------------------------------------------------------
    # Берём ВСЕ ссылки внутри дерева.
    #
    # Поэтому автоматически получаем:
    # категория
    # └── подкатегория
    #     └── подподкатегория
    #         └── ...
    # --------------------------------------------------------

    for a in block.select(
        "ul.tree a[href]"
    ):

        href = a.get(
            "href",
            ""
        )

        href = absolute_url(
            href
        )

        href = base_product_url(
            href
        )

        if not href:
            continue

        if DOMAIN not in href:
            continue

        name = clean(
            a.get_text(
                " ",
                strip=True
            )
        )

        if not name:
            continue

        if href in seen:
            continue

        seen.add(href)

        categories.append({
            "name": name,
            "url": href
        })

    if CATEGORY_LIMIT:
        categories = categories[
            :CATEGORY_LIMIT
        ]

    print(
        f"📂 Найдено категорий: "
        f"{len(categories)}"
    )

    for i, category in enumerate(
        categories,
        1
    ):

        print(
            f"   {i}. "
            f"{category['name']} → "
            f"{category['url']}"
        )

    return categories


# ============================================================
# SHOW ALL
# ============================================================

def get_show_all_data(soup):

    form = soup.select_one(
        "form.showall"
    )

    if not form:
        return None

    action = form.get(
        "action",
        ""
    )

    action = absolute_url(
        action
    )

    n_input = form.select_one(
        "input[name='n']"
    )

    category_input = form.select_one(
        "input[name='id_category']"
    )

    if not n_input:
        return None

    n = clean(
        n_input.get(
            "value",
            ""
        )
    )

    if not n:
        return None

    data = {
        "n": n
    }

    if category_input:

        category_id = clean(
            category_input.get(
                "value",
                ""
            )
        )

        if category_id:
            data["id_category"] = category_id

    return {
        "url": action,
        "params": data
    }


# ============================================================
# ТОВАРЫ В КАТЕГОРИИ
# ============================================================

def extract_product_links(soup):

    links = []

    seen = set()

    selectors = [

        "ul.product_list "
        "li.ajax_block_product "
        "a.product-name",

        "ul.product_list "
        "li.ajax_block_product "
        "h5 a",

        "ul.product_list "
        "li.ajax_block_product "
        "a[href*='.html']",

        ".product-container "
        "a.product-name",

        ".product-container "
        "h5 a",

        ".product-container "
        "a[href*='.html']",

        "li.product-box.item "
        "a.product-image",

        "li.product-box.item "
        "h5.product-name a",

        "li.product-box.item "
        "a[href*='.html']",

        "#productscategory_list "
        "li.product-box "
        "a[href*='.html']",
    ]

    for selector in selectors:

        for a in soup.select(
            selector
        ):

            href = a.get(
                "href",
                ""
            )

            if not href:
                continue

            href = absolute_url(
                href
            )

            href = base_product_url(
                href
            )

            if not href:
                continue

            if DOMAIN not in href:
                continue

            if ".html" not in href:
                continue

            # ------------------------------------------------
            # Отбрасываем явно не товарные ссылки
            # ------------------------------------------------

            if href in seen:
                continue

            seen.add(href)

            links.append(href)

    return links


# ============================================================
# ПАГИНАЦИЯ
# ============================================================

def get_page_numbers(soup):

    pages = []

    pagination = soup.select_one(
        "#pagination_bottom"
    )

    if not pagination:
        return pages

    for a in pagination.select(
        "a[href]"
    ):

        href = a.get(
            "href",
            ""
        )

        match = re.search(
            r"#/page-(\d+)",
            href
        )

        if match:

            page = int(
                match.group(1)
            )

            if page not in pages:
                pages.append(page)

    pages.sort()

    return pages


# ============================================================
# КАТЕГОРИЯ
# ============================================================

def parse_category(category_url):

    print()
    print(
        f"📁 КАТЕГОРИЯ: {category_url}"
    )

    soup = get_soup(
        category_url
    )

    if not soup:
        return []

    # --------------------------------------------------------
    # Сначала пытаемся Показать все
    # --------------------------------------------------------

    show_all = get_show_all_data(
        soup
    )

    if show_all:

        print(
            f"   🔎 Пытаемся загрузить "
            f"все товары сразу: "
            f"n={show_all['params'].get('n')}"
        )

        all_soup = get_soup(
            show_all["url"],
            params=show_all["params"]
        )

        if all_soup:

            links = extract_product_links(
                all_soup
            )

            if links:

                print(
                    f"   ✅ Получено товаров: "
                    f"{len(links)}"
                )

                return links

            else:

                print(
                    "   ⚠️ Показать все "
                    "не вернуло товары"
                )

    # --------------------------------------------------------
    # Обычная первая страница
    # --------------------------------------------------------

    links = extract_product_links(
        soup
    )

    print(
        f"   📦 Товаров на первой странице: "
        f"{len(links)}"
    )

    all_links = list(
        dict.fromkeys(links)
    )

    # --------------------------------------------------------
    # Получаем номера страниц
    # --------------------------------------------------------

    pages = get_page_numbers(
        soup
    )

    if pages:

        print(
            f"   📄 Дополнительные страницы: "
            f"{pages}"
        )

    # --------------------------------------------------------
    # Если Show All не сработал,
    # пробуем ?page=N
    # --------------------------------------------------------

    for page in pages:

        if page <= 1:
            continue

        # В исходном сайте ссылки выглядят как:
        # #/page-2
        #
        # requests fragment не отправляет,
        # поэтому пробуем обычный параметр page.
        page_soup = get_soup(
            category_url,
            params={
                "page": page
            }
        )

        if not page_soup:
            continue

        page_links = extract_product_links(
            page_soup
        )

        print(
            f"   📄 Страница {page}: "
            f"{len(page_links)} товаров"
        )

        for link in page_links:

            if link not in all_links:
                all_links.append(link)

    print(
        f"   ✅ Всего товаров в категории: "
        f"{len(all_links)}"
    )

    return all_links


# ============================================================
# ТОВАР — TITLE
# ============================================================

def parse_title(soup):

    element = soup.select_one(
        "h1[itemprop='name']"
    )

    if element:

        return clean(
            element.get_text(
                " ",
                strip=True
            )
        )

    return ""


# ============================================================
# ТОВАР — SKU
# ============================================================

def parse_sku(soup):

    element = soup.select_one(
        "#product_reference "
        "span[itemprop='sku']"
    )

    if element:

        sku = element.get(
            "content"
        )

        if sku:
            return clean(sku)

        return clean(
            element.get_text(
                " ",
                strip=True
            )
        )

    # fallback
    element = soup.select_one(
        "#product_reference .editable"
    )

    if element:

        sku = element.get(
            "content"
        )

        if sku:
            return clean(sku)

        return clean(
            element.get_text(
                " ",
                strip=True
            )
        )

    return ""


# ============================================================
# ТОВАР — STATUS
# ============================================================

def parse_status(soup):

    # --------------------------------------------------------
    # Сначала видимый текст
    # --------------------------------------------------------

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

        if text:

            text_lower = text.lower()

            if (
                "відсут" in text_lower
                or
                "отсутств" in text_lower
                or
                "немає" in text_lower
                or
                "нет" in text_lower
            ):

                return "Немає в наявності"

            return text

    # --------------------------------------------------------
    # Schema.org
    # --------------------------------------------------------

    availability = soup.select_one(
        "link[itemprop='availability']"
    )

    if availability:

        href = clean(
            availability.get(
                "href",
                ""
            )
        ).lower()

        if "instock" in href:

            return "В наявності"

        if "outofstock" in href:

            return "Немає в наявності"

    # --------------------------------------------------------
    # Если ничего не нашли
    # --------------------------------------------------------

    return ""


# ============================================================
# ТОВАР — ЦВЕТА
# ============================================================

def parse_colors(soup):

    colors = []

    attributes = soup.select(
        "#attributes "
        ".attribute_fieldset"
    )

    for fieldset in attributes:

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

        # Нас интересует именно Колір
        if "колір" not in label_text:
            continue

        for li in fieldset.select(
            ".attribute_list li"
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

            # ----------------------------------------------
            # Получаем название цвета
            # ----------------------------------------------

            color_name = ""

            spans = li.select(
                "span"
            )

            for span in spans:

                text = clean(
                    span.get_text(
                        " ",
                        strip=True
                    )
                )

                if text:

                    # Не берём span checked
                    if text.lower() not in (
                        "checked",
                    ):

                        color_name = text

            # ----------------------------------------------
            # Если span не нашли
            # ----------------------------------------------

            if not color_name:

                text = clean(
                    li.get_text(
                        " ",
                        strip=True
                    )
                )

                color_name = text

            if not color_name:
                continue

            colors.append({
                "id": color_id,
                "name": color_name,
                "slug": translit_slug(
                    color_name
                )
            })

    # --------------------------------------------------------
    # Удаляем дубликаты
    # --------------------------------------------------------

    result = []
    seen = set()

    for color in colors:

        key = (
            color["id"],
            color["name"].lower()
        )

        if key in seen:
            continue

        seen.add(key)

        result.append(color)

    return result


# ============================================================
# URL ЦВЕТА
# ============================================================

def make_color_url(
    product_url,
    color
):

    product_url = base_product_url(
        product_url
    )

    color_id = color.get(
        "id",
        ""
    )

    slug = color.get(
        "slug",
        ""
    )

    if not color_id or not slug:
        return product_url

    return (
        f"{product_url}"
        f"#/{color_id}-kolir-{slug}"
    )


# ============================================================
# ТОВАР
# ============================================================

def parse_product(
    url,
    price_map
):

    url = base_product_url(
        url
    )

    soup = get_soup(
        url
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

    # --------------------------------------------------------
    # ЦЕНА ИЗ GOOGLE SHEETS
    # --------------------------------------------------------

    price = price_map.get(
        normalize_sku(sku)
    )

    if price is None:
        price = ""

    colors = parse_colors(
        soup
    )

    rows = []

    # --------------------------------------------------------
    # Если цветов нет
    # --------------------------------------------------------

    if not colors:

        rows.append({
            "SKU": sku,
            "TITLE": title,
            "PRICE USD": price,
            "STATUS": status,
            "COLOR": "",
            "URL": url
        })

        return rows

    # --------------------------------------------------------
    # Если есть цвета
    #
    # Каждый цвет = отдельная строка
    # --------------------------------------------------------

    for color in colors:

        color_url = make_color_url(
            url,
            color
        )

        rows.append({
            "SKU": sku,
            "TITLE": title,
            "PRICE USD": price,
            "STATUS": status,
            "COLOR": color["name"],
            "URL": color_url
        })

    return rows


# ============================================================
# EXCEL
# ============================================================

def save_excel(rows):

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

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

    for row in rows:

        ws.append([
            row.get(
                "SKU",
                ""
            ),
            row.get(
                "TITLE",
                ""
            ),
            row.get(
                "PRICE USD",
                ""
            ),
            row.get(
                "STATUS",
                ""
            ),
            row.get(
                "COLOR",
                ""
            ),
            row.get(
                "URL",
                ""
            )
        ])

    # --------------------------------------------------------
    # Формат цены
    # --------------------------------------------------------

    price_column = 3

    for cell in ws.iter_cols(
        min_col=price_column,
        max_col=price_column,
        min_row=2
    ):

        for c in cell:

            if isinstance(
                c.value,
                (int, float)
            ):

                c.number_format = (
                    '0.00'
                )

    # --------------------------------------------------------
    # Автоширина
    # --------------------------------------------------------

    widths = {
        "A": 20,
        "B": 60,
        "C": 15,
        "D": 25,
        "E": 20,
        "F": 100
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    # --------------------------------------------------------
    # Закрепляем шапку
    # --------------------------------------------------------

    ws.freeze_panes = "A2"

    wb.save(
        FILE_PATH
    )

    print()
    print(
        f"💾 Excel сохранён:"
    )

    print(
        FILE_PATH
    )


# ============================================================
# RUN PARSER
# ============================================================

def run_parser():

    print()
    print("🔥 BOT STARTED")
    print("🚀 BOT RUNNING")
    print("🚀 STARTED ArizoneSports")
    print(
        f"📄 SCRIPT: parsers/ArizoneSports/run.py"
    )

    if not create_lock():

        return

    save_status(
        "running",
        0,
        0,
        "Загрузка Google Sheets"
    )

    try:

        # ====================================================
        # 1. Загружаем цены
        # ====================================================

        price_map = (
            get_google_sheet_price_map()
        )

        # ====================================================
        # 2. Получаем категории
        # ====================================================

        save_status(
            "running",
            0,
            0,
            "Получение категорий"
        )

        categories = get_categories()

        if not categories:

            print(
                "❌ Категории не найдены"
            )

            save_status(
                "error",
                0,
                0,
                "Категории не найдены"
            )

            return

        # ====================================================
        # 3. Обходим все категории
        # ====================================================

        all_product_links = []

        seen_products = set()

        for category_index, category in enumerate(
            categories,
            1
        ):

            print()
            print(
                f"📂 [{category_index}/"
                f"{len(categories)}] "
                f"{category['name']}"
            )

            save_status(
                "running",
                category_index,
                len(categories),
                category["name"]
            )

            product_links = parse_category(
                category["url"]
            )

            for product_url in product_links:

                product_url = base_product_url(
                    product_url
                )

                if product_url in seen_products:
                    continue

                seen_products.add(
                    product_url
                )

                all_product_links.append(
                    product_url
                )

        print()
        print(
            f"🛒 ВСЕГО УНИКАЛЬНЫХ ТОВАРОВ: "
            f"{len(all_product_links)}"
        )

        # ====================================================
        # 4. Ограничение для теста
        # ====================================================

        if PRODUCT_LIMIT:

            all_product_links = (
                all_product_links[
                    :PRODUCT_LIMIT
                ]
            )

            print(
                f"⚠️ PRODUCT_LIMIT = "
                f"{PRODUCT_LIMIT}"
            )

        # ====================================================
        # 5. Парсим товары
        # ====================================================

        result_rows = []

        total = len(
            all_product_links
        )

        for index, product_url in enumerate(
            all_product_links,
            1
        ):

            print(
                f"📦 [{index}/{total}] "
                f"{product_url}"
            )

            save_status(
                "running",
                index,
                total,
                product_url
            )

            try:

                product_rows = parse_product(
                    product_url,
                    price_map
                )

                for row in product_rows:

                    result_rows.append(
                        row
                    )

                if product_rows:

                    first = product_rows[0]

                    price_text = (
                        first["PRICE USD"]
                        if first["PRICE USD"] != ""
                        else "НЕТ ЦЕНЫ"
                    )

                    print(
                        f"   ✅ "
                        f"{first['SKU']} | "
                        f"{price_text} USD | "
                        f"{len(product_rows)} "
                        f"цветов"
                    )

                else:

                    print(
                        "   ⚠️ Товар не разобран"
                    )

            except Exception as e:

                print(
                    f"   ❌ Ошибка товара: "
                    f"{e}"
                )

            # Небольшая пауза
            time.sleep(0.15)

        # ====================================================
        # 6. Дедупликация
        #
        # SKU + COLOR
        #
        # Поэтому цвета НЕ объединяются.
        # ====================================================

        final_rows = []

        seen_rows = set()

        for row in result_rows:

            sku = normalize_sku(
                row.get(
                    "SKU",
                    ""
                )
            )

            color = clean(
                row.get(
                    "COLOR",
                    ""
                )
            ).lower()

            url = base_product_url(
                row.get(
                    "URL",
                    ""
                )
            )

            if sku:

                key = (
                    sku,
                    color
                )

            else:

                key = (
                    url,
                    color
                )

            if key in seen_rows:
                continue

            seen_rows.add(
                key
            )

            final_rows.append(
                row
            )

        # ====================================================
        # 7. Сохраняем Excel
        # ====================================================

        save_excel(
            final_rows
        )

        # ====================================================
        # 8. Финальный статус
        # ====================================================

        save_status(
            "done",
            len(final_rows),
            len(final_rows),
            "Парсер завершён"
        )

        print()
        print("=" * 60)
        print("🎉 ПАРСЕР ЗАВЕРШЁН")
        print(
            f"📦 Строк в Excel: "
            f"{len(final_rows)}"
        )
        print(
            f"💰 SKU с ценами: "
            f"{sum("
                "1 for row in final_rows "
                "if row.get('PRICE USD') != ''"
            )}"
        )
        print(
            f"📄 Файл: "
            f"{FILE_PATH}"
        )
        print("=" * 60)

    except Exception as e:

        print()
        print(
            f"💥 КРИТИЧЕСКАЯ ОШИБКА: {e}"
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
