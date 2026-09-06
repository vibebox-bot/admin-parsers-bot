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
# USER
# =========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "default"


# =========================================================
# SETTINGS
# =========================================================

BASE = "https://arizonesports.com.ua/uk"

# None = все категории
# Для теста можно поставить, например: 2
CATEGORY_LIMIT = 1

# None = все товары
# Для теста можно поставить, например: 20
PRODUCT_LIMIT = None


# =========================================================
# OUTPUT
# =========================================================

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


# =========================================================
# HTTP
# =========================================================

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


# =========================================================
# LOCK
# =========================================================

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
                print("⚠️ Parser already running")
                return False

        except Exception:
            pass

    try:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))

        return True

    except Exception as e:
        print(f"❌ Cannot create lock: {e}")
        return False


def remove_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


# =========================================================
# STATUS
# =========================================================

def save_status(
    status,
    message="",
    current=0,
    total=0,
    category=""
):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "status": status,
        "message": message,
        "current": current,
        "total": total,
        "category": category,
        "updated_at": datetime.now().isoformat()
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


# =========================================================
# HELPERS
# =========================================================

def clean(value):
    if value is None:
        return ""

    value = str(value)

    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)

    return value.strip()


def absolute_url(url):
    if not url:
        return ""

    url = url.strip()

    # Убираем fragment #/43-kolir...
    # Цветовые варианты НЕ являются отдельными товарами
    url = url.split("#", 1)[0]

    if url.startswith("http://") or url.startswith("https://"):
        return url

    if url.startswith("/"):
        return "https://arizonesports.com.ua" + url

    return BASE.rstrip("/") + "/" + url.lstrip("/")


# =========================================================
# GET SOUP
# =========================================================

def get_soup(url, params=None, retries=3):
    for attempt in range(1, retries + 1):

        try:
            response = session.get(
                url,
                params=params,
                timeout=40
            )

            response.raise_for_status()

            return BeautifulSoup(
                response.text,
                "html.parser"
            )

        except Exception as e:

            print(
                f"⚠️ Request error "
                f"{attempt}/{retries}: {url}"
            )
            print(e)

            if attempt < retries:
                time.sleep(2)

    return None


# =========================================================
# LOGIN
# =========================================================
#
# На ArizoneSports авторизация пока не используется.
# Оставляем функцию в структуре парсера.
# =========================================================

def login():
    print("🔐 Login not required")
    return True


# =========================================================
# CATEGORIES
# =========================================================

def get_categories():
    """
    Забираем ВСЕ категории и подкатегории
    из блока #categories_block_left.

    Не ограничиваемся первым уровнем.
    Любая глубина дерева будет обработана.
    """

    print("📂 Получаем категории...")

    soup = get_soup(BASE)

    if not soup:
        print("❌ Не удалось получить главную страницу")
        return []

    categories_block = soup.select_one(
        "#categories_block_left"
    )

    if not categories_block:
        print(
            "❌ Не найден #categories_block_left"
        )
        return []

    categories = []
    seen = set()

    # Берём все ссылки внутри дерева
    for link in categories_block.select(
        "ul.tree a[href]"
    ):

        href = link.get("href")

        if not href:
            continue

        url = absolute_url(href)

        if not url:
            continue

        # Только наш сайт
        if "arizonesports.com.ua" not in url:
            continue

        # Не добавляем дубли
        if url in seen:
            continue

        name = clean(
            link.get_text(" ", strip=True)
        )

        if not name:
            continue

        seen.add(url)

        categories.append({
            "name": name,
            "url": url
        })

    print(
        f"📂 Найдено категорий/подкатегорий: "
        f"{len(categories)}"
    )

    for i, category in enumerate(
        categories,
        start=1
    ):
        print(
            f"   {i}. "
            f"{category['name']} -> "
            f"{category['url']}"
        )

    return categories


# =========================================================
# SHOW ALL
# =========================================================

def get_show_all_url(category_url, soup):
    """
    На странице категории есть форма:

    <form class="showall" ...>
        <input name="n" value="85">
        <input name="id_category" value="14">
    </form>

    Используем её, чтобы получить все товары
    категории одним запросом.
    """

    form = soup.select_one(
        "form.showall"
    )

    if not form:
        return None, None

    action = form.get("action") or category_url

    action = absolute_url(action)

    params = {}

    for inp in form.select(
        "input[name]"
    ):
        name = inp.get("name")
        value = inp.get("value", "")

        if name:
            params[name] = value

    # Обычно здесь есть n=85
    n = params.get("n")

    if not n:
        return None, None

    try:
        n_int = int(n)
    except Exception:
        return None, None

    if n_int <= 0:
        return None, None

    return action, params


# =========================================================
# PRODUCT LINKS
# =========================================================

def extract_product_links(soup):
    """
    Ищем ссылки на товары.

    Важно:
    #/43-kolir-bezhevij
    #/26-kolir-zelenij

    отбрасываются, потому что это варианты
    одного товара.
    """

    links = []
    seen = set()

    selectors = [
        "ul.product_list li.ajax_block_product h5 a[href]",
        "ul.product_list li.ajax_block_product a.product-name[href]",
        "ul.product_list li.ajax_block_product a[href]",
        ".product-container a.product-name[href]",
        ".product-container h5 a[href]",
        ".product-container a[href]",
    ]

    for selector in selectors:

        for link in soup.select(selector):

            href = link.get("href")

            if not href:
                continue

            url = absolute_url(href)

            if not url:
                continue

            # Только товары нашего сайта
            if "arizonesports.com.ua" not in url:
                continue

            # Нам нужны HTML-страницы товаров
            if ".html" not in url.lower():
                continue

            # Убираем fragment
            url = url.split("#", 1)[0]

            if url in seen:
                continue

            seen.add(url)
            links.append(url)

    return links


# =========================================================
# CATEGORY PARSER
# =========================================================

def parse_category(category):
    """
    Получаем все товары из одной категории.

    Сначала пробуем "Показати все".

    Если по какой-то причине оно не сработало,
    используем обычную страницу категории.
    """

    category_name = category["name"]
    category_url = category["url"]

    print()
    print("=" * 80)
    print(f"📁 CATEGORY: {category_name}")
    print(category_url)
    print("=" * 80)

    save_status(
        "category",
        f"Обработка категории: {category_name}",
        category=category_name
    )

    soup = get_soup(category_url)

    if not soup:
        print(
            "❌ Не удалось получить категорию"
        )
        return []

    # -----------------------------------------------------
    # 1. SHOW ALL
    # -----------------------------------------------------

    showall_url, showall_params = get_show_all_url(
        category_url,
        soup
    )

    if showall_url:

        print(
            f"📄 Найдена кнопка 'Показати все': "
            f"{showall_params}"
        )

        all_soup = get_soup(
            showall_url,
            params=showall_params
        )

        if all_soup:

            product_links = extract_product_links(
                all_soup
            )

            if product_links:

                print(
                    f"✅ Показати все: "
                    f"{len(product_links)} товаров"
                )

                return product_links

            print(
                "⚠️ 'Показати все' открылось, "
                "но товары не найдены"
            )

    # -----------------------------------------------------
    # 2. CURRENT PAGE
    # -----------------------------------------------------

    product_links = extract_product_links(
        soup
    )

    print(
        f"📦 На текущей странице: "
        f"{len(product_links)} товаров"
    )

    # -----------------------------------------------------
    # 3. FALLBACK PAGINATION
    # -----------------------------------------------------
    #
    # На сайте ссылки имеют:
    #
    # /category#/page-2
    #
    # Fragment сервер не получает.
    #
    # Поэтому сначала пытаемся найти
    # реальные ссылки пагинации / параметры.
    # -----------------------------------------------------

    pagination = soup.select_one(
        "#pagination_bottom"
    )

    if not pagination:
        pagination = soup.select_one(
            ".pagination"
        )

    if pagination:

        page_numbers = set()

        for a in pagination.select("a[href]"):

            href = a.get("href", "")

            # Ищем #/page-2
            match = re.search(
                r"#/page-(\d+)",
                href
            )

            if match:
                try:
                    page_numbers.add(
                        int(match.group(1))
                    )
                except Exception:
                    pass

            # На случай обычной ?page=2
            match = re.search(
                r"[?&]page=(\d+)",
                href
            )

            if match:
                try:
                    page_numbers.add(
                        int(match.group(1))
                    )
                except Exception:
                    pass

        if page_numbers:

            print(
                f"📄 Найдены страницы: "
                f"{sorted(page_numbers)}"
            )

            for page in sorted(page_numbers):

                # Сначала пробуем ?page=N
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

                if page_links:

                    print(
                        f"   📄 Страница {page}: "
                        f"{len(page_links)} товаров"
                    )

                    for url in page_links:

                        if url not in product_links:
                            product_links.append(url)

    # -----------------------------------------------------
    # DEDUPE
    # -----------------------------------------------------

    unique = []
    seen = set()

    for url in product_links:

        url = absolute_url(url)

        if url in seen:
            continue

        seen.add(url)
        unique.append(url)

    print(
        f"📦 Итого товаров в категории: "
        f"{len(unique)}"
    )

    return unique


# =========================================================
# STATUS
# =========================================================

def parse_status(soup):
    """
    Определяем наличие товара.

    Вариант 1:
    #availability_value

    Например:
    Наразі товар відсутній

    Вариант 2:
    schema.org:

    <link itemprop="availability"
          href="https://schema.org/InStock">

    """

    # -----------------------------------------------------
    # 1. availability_value
    # -----------------------------------------------------

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

            lower = text.lower()

            if (
                "відсут" in lower
                or "немає" in lower
                or "нема" in lower
                or "out of stock" in lower
            ):
                return "Немає в наявності"

            if (
                "наяв" in lower
                or "in stock" in lower
            ):
                return "В наявності"

            return text

    # -----------------------------------------------------
    # 2. Schema availability
    # -----------------------------------------------------

    availability_link = soup.select_one(
        "[itemprop='availability']"
    )

    if availability_link:

        href = (
            availability_link.get("href")
            or ""
        ).lower()

        if "instock" in href:
            return "В наявності"

        if (
            "outofstock" in href
            or "soldout" in href
        ):
            return "Немає в наявності"

    # -----------------------------------------------------
    # 3. Last quantities
    # -----------------------------------------------------

    last_quantities = soup.select_one(
        "#last_quantities"
    )

    if last_quantities:

        text = clean(
            last_quantities.get_text(
                " ",
                strip=True
            )
        )

        if text:
            return "В наявності"

    # -----------------------------------------------------
    # 4. Default
    # -----------------------------------------------------

    return "Невідомо"


# =========================================================
# COLORS
# =========================================================

def parse_colors(soup):
    """
    Получаем все доступные цвета товара.

    Например:

    Блакитний
    Зелений
    Рожевий
    Чорний
    Фіолетовий

    Важно:

    URL:
    #/43-kolir-bezhevij
    #/26-kolir-zelenij

    НЕ обрабатываем отдельно.

    Это варианты одного товара.
    Все варианты собираем из #attributes.
    """

    colors = []

    # -----------------------------------------------------
    # Ищем все группы атрибутов
    # -----------------------------------------------------

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

        # Нас интересует только Колір
        if "колір" not in label_text:
            continue

        # -------------------------------------------------
        # Варианты цветов
        # -------------------------------------------------

        for li in fieldset.select(
            ".attribute_list li"
        ):

            # Пытаемся взять span с названием
            color_spans = li.select(
                "span"
            )

            color = ""

            for span in color_spans:

                text = clean(
                    span.get_text(
                        " ",
                        strip=True
                    )
                )

                if not text:
                    continue

                # Пропускаем технические элементы
                if text.lower() in {
                    "checked",
                    "radio"
                }:
                    continue

                # Если внутри есть название цвета
                color = text

            # Если span не дал результат
            if not color:
                color = clean(
                    li.get_text(
                        " ",
                        strip=True
                    )
                )

            if not color:
                continue

            # Убираем технический мусор
            color = re.sub(
                r"\s+",
                " ",
                color
            ).strip()

            # Иногда текст может содержать
            # лишние служебные слова
            color = color.replace(
                "checked",
                ""
            ).strip()

            if (
                color
                and color not in colors
            ):
                colors.append(color)

    # -----------------------------------------------------
    # Альтернативный поиск,
    # если структура немного отличается
    # -----------------------------------------------------

    if not colors:

        attributes = soup.select_one(
            "#attributes"
        )

        if attributes:

            for li in attributes.select(
                "li"
            ):

                text = clean(
                    li.get_text(
                        " ",
                        strip=True
                    )
                )

                if not text:
                    continue

                # Проверяем, что это часть
                # группы Колір
                parent_fieldset = li.find_parent(
                    "fieldset"
                )

                if not parent_fieldset:
                    continue

                label = parent_fieldset.select_one(
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

                if "колір" not in label_text:
                    continue

                if text not in colors:
                    colors.append(text)

    return ", ".join(colors)


# =========================================================
# SKU / ARTICLE
# =========================================================

def parse_sku(soup):
    """
    Основной источник:

    <span
        class="editable"
        itemprop="sku"
        content="BBTT-Q7">
        BBTT-Q7
    </span>

    Сначала берём content,
    затем текст,
    затем запасной вариант.
    """

    # -----------------------------------------------------
    # 1. itemprop=sku
    # -----------------------------------------------------

    sku_el = soup.select_one(
        "#product_reference [itemprop='sku']"
    )

    if sku_el:

        sku = clean(
            sku_el.get("content")
            or sku_el.get_text(
                " ",
                strip=True
            )
        )

        if sku:
            return sku

    # -----------------------------------------------------
    # 2. .editable
    # -----------------------------------------------------

    sku_el = soup.select_one(
        "#product_reference .editable"
    )

    if sku_el:

        sku = clean(
            sku_el.get("content")
            or sku_el.get_text(
                " ",
                strip=True
            )
        )

        if sku:
            return sku

    # -----------------------------------------------------
    # 3. product_reference text
    # -----------------------------------------------------

    ref = soup.select_one(
        "#product_reference"
    )

    if ref:

        text = clean(
            ref.get_text(
                " ",
                strip=True
            )
        )

        text = re.sub(
            r"^Артикул\s*",
            "",
            text,
            flags=re.I
        )

        text = clean(text)

        if text:
            return text

    return ""


# =========================================================
# PRODUCT
# =========================================================

def parse_product(url):
    """
    Парсим одну карточку товара.
    """

    print(
        f"      🔎 {url}"
    )

    soup = get_soup(url)

    if not soup:
        print(
            "      ❌ Не удалось открыть товар"
        )
        return None

    # -----------------------------------------------------
    # TITLE
    # -----------------------------------------------------

    title_el = soup.select_one(
        "h1[itemprop='name']"
    )

    if not title_el:
        title_el = soup.select_one(
            "h1"
        )

    title = ""

    if title_el:
        title = clean(
            title_el.get_text(
                " ",
                strip=True
            )
        )

    # -----------------------------------------------------
    # SKU
    # -----------------------------------------------------

    sku = parse_sku(
        soup
    )

    # -----------------------------------------------------
    # STATUS
    # -----------------------------------------------------

    status = parse_status(
        soup
    )

    # -----------------------------------------------------
    # COLORS
    # -----------------------------------------------------

    colors = parse_colors(
        soup
    )

    # -----------------------------------------------------
    # PRICE USD
    # -----------------------------------------------------
    #
    # Пока НЕ берём цену с сайта.
    #
    # Здесь позже подключим Google Sheet
    # и будем искать цену по SKU.
    # -----------------------------------------------------

    price_usd = ""

    # -----------------------------------------------------
    # RESULT
    # -----------------------------------------------------

    product = {
        "SKU": sku,
        "TITLE": title,
        "PRICE USD": price_usd,
        "STATUS": status,
        "COLORS": colors,
        "URL": url
    }

    print(
        f"      ✅ SKU: {sku or '—'}"
    )

    print(
        f"      📦 STATUS: {status}"
    )

    print(
        f"      🎨 COLORS: "
        f"{colors or '—'}"
    )

    return product


# =========================================================
# SAVE EXCEL
# =========================================================

def save_excel(products):
    """
    Сохраняем результат в Excel.
    """

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
        "COLORS",
        "URL"
    ]

    ws.append(headers)

    for product in products:

        ws.append([
            product.get("SKU", ""),
            product.get("TITLE", ""),
            product.get("PRICE USD", ""),
            product.get("STATUS", ""),
            product.get("COLORS", ""),
            product.get("URL", "")
        ])

    # Автоширина
    widths = {
        "A": 20,
        "B": 60,
        "C": 15,
        "D": 25,
        "E": 40,
        "F": 90
    }

    for column, width in widths.items():
        ws.column_dimensions[
            column
        ].width = width

    # Freeze header
    ws.freeze_panes = "A2"

    # Filter
    ws.auto_filter.ref = (
        ws.dimensions
    )

    # Временный файл
    temp_file = FILE_PATH + ".tmp"

    wb.save(
        temp_file
    )

    # Замена старого
    if os.path.exists(FILE_PATH):
        os.remove(FILE_PATH)

    os.replace(
        temp_file,
        FILE_PATH
    )

    print()
    print(
        f"💾 Excel сохранён:"
    )
    print(
        FILE_PATH
    )


# =========================================================
# RUN PARSER
# =========================================================

def run_parser():

    if not create_lock():
        return

    start_time = time.time()

    try:

        print()
        print("🔥 ARIZONE SPORTS PARSER")
        print("=" * 80)
        print(
            f"🌐 BASE: {BASE}"
        )
        print(
            f"📁 OUTPUT: {FILE_PATH}"
        )
        print("=" * 80)

        save_status(
            "starting",
            "Запуск парсера"
        )

        # -------------------------------------------------
        # LOGIN
        # -------------------------------------------------

        if not login():

            print(
                "❌ Login failed"
            )

            save_status(
                "error",
                "Login failed"
            )

            return

        # -------------------------------------------------
        # CATEGORIES
        # -------------------------------------------------

        categories = get_categories()

        if not categories:

            print(
                "❌ Категории не найдены"
            )

            save_status(
                "error",
                "Категории не найдены"
            )

            return

        # -------------------------------------------------
        # CATEGORY LIMIT
        # -------------------------------------------------

        if CATEGORY_LIMIT is not None:

            categories = categories[
                :CATEGORY_LIMIT
            ]

            print()
            print(
                f"🧪 TEST MODE: "
                f"{len(categories)} категорий"
            )

        # -------------------------------------------------
        # COLLECT PRODUCT URLS
        # -------------------------------------------------

        all_product_urls = []

        seen_urls = set()

        total_categories = len(
            categories
        )

        for category_index, category in enumerate(
            categories,
            start=1
        ):

            print()
            print(
                f"📂 CATEGORY "
                f"{category_index}/"
                f"{total_categories}"
            )

            save_status(
                "category",
                f"{category_index}/{total_categories}",
                current=category_index,
                total=total_categories,
                category=category["name"]
            )

            product_urls = parse_category(
                category
            )

            for product_url in product_urls:

                product_url = absolute_url(
                    product_url
                )

                if not product_url:
                    continue

                if product_url in seen_urls:
                    continue

                seen_urls.add(
                    product_url
                )

                all_product_urls.append(
                    product_url
                )

            print(
                f"📊 Всего уникальных товаров "
                f"после категории: "
                f"{len(all_product_urls)}"
            )

            # -------------------------------------------------
            # PRODUCT LIMIT
            # -------------------------------------------------

            if (
                PRODUCT_LIMIT is not None
                and len(all_product_urls)
                >= PRODUCT_LIMIT
            ):

                all_product_urls = (
                    all_product_urls[
                        :PRODUCT_LIMIT
                    ]
                )

                print(
                    f"🧪 PRODUCT LIMIT: "
                    f"{PRODUCT_LIMIT}"
                )

                break

        # -------------------------------------------------
        # NO PRODUCTS
        # -------------------------------------------------

        if not all_product_urls:

            print()
            print(
                "❌ Товары не найдены"
            )

            save_status(
                "error",
                "Товары не найдены"
            )

            return

        print()
        print("=" * 80)
        print(
            f"🛒 ВСЕГО УНИКАЛЬНЫХ ТОВАРОВ: "
            f"{len(all_product_urls)}"
        )
        print("=" * 80)

        # -------------------------------------------------
        # PARSE PRODUCTS
        # -------------------------------------------------

        products = []

        seen_skus = set()

        total_products = len(
            all_product_urls
        )

        for index, product_url in enumerate(
            all_product_urls,
            start=1
        ):

            print()
            print(
                f"📦 PRODUCT "
                f"{index}/{total_products}"
            )

            save_status(
                "product",
                f"{index}/{total_products}",
                current=index,
                total=total_products
            )

            product = parse_product(
                product_url
            )

            if not product:
                continue

            sku = clean(
                product.get("SKU", "")
            )

            # -------------------------------------------------
            # DEDUPE BY SKU
            # -------------------------------------------------
            #
            # Если SKU есть — считаем его главным
            # идентификатором товара.
            #
            # Цветовые #/ варианты сюда не попадут,
            # потому что URL уже очищается от fragment.
            # -------------------------------------------------

            if sku:

                sku_key = sku.lower()

                if sku_key in seen_skus:

                    print(
                        f"      ⚠️ Дубликат SKU: "
                        f"{sku}"
                    )

                    continue

                seen_skus.add(
                    sku_key
                )

            products.append(
                product
            )

        # -------------------------------------------------
        # SAVE
        # -------------------------------------------------

        print()
        print("=" * 80)
        print(
            f"💾 Сохраняем "
            f"{len(products)} товаров..."
        )
        print("=" * 80)

        save_excel(
            products
        )

        # -------------------------------------------------
        # FINISH
        # -------------------------------------------------

        elapsed = (
            time.time()
            - start_time
        )

        print()
        print("=" * 80)
        print("✅ PARSER FINISHED")
        print(
            f"📦 Товаров: {len(products)}"
        )
        print(
            f"⏱ Время: "
            f"{elapsed:.1f} сек."
        )
        print(
            f"📄 FILE: {FILE_PATH}"
        )
        print("=" * 80)

        save_status(
            "done",
            f"Готово. Товаров: {len(products)}",
            current=len(products),
            total=len(products)
        )

    except Exception as e:

        print()
        print("=" * 80)
        print("❌ PARSER ERROR")
        print("=" * 80)

        print(
            type(e).__name__,
            str(e)
        )

        save_status(
            "error",
            f"{type(e).__name__}: {e}"
        )

        raise

    finally:

        remove_lock()


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    run_parser()
