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

print("🔥 ARIZONE_Sports")

BASE = "https://arizonesports.com.ua/uk"


# =========================
# ⚙️ SWITCH
# =========================

# CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

# Для теста можно ограничить количество товаров.
# Например:
# PRODUCT_LIMIT = 20
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

def get_soup(url, params=None):

    for attempt in range(3):

        try:

            r = session.get(
                url,
                params=params,
                timeout=30
            )

            if r.status_code == 200:

                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

            print(
                f"⚠ HTTP {r.status_code}: {url}"
            )

        except Exception as e:

            print(
                f"⚠ Ошибка запроса "
                f"{attempt + 1}/3: {url}"
            )

        time.sleep(1)

    return BeautifulSoup(
        "",
        "html.parser"
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

def absolute_url(url):

    if not url:
        return ""

    url = url.strip()

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("/"):
        return "https://arizonesports.com.ua" + url

    if url.startswith("http://"):
        return url.replace(
            "http://",
            "https://",
            1
        )

    if url.startswith("https://"):
        return url

    return url


# =========================
# CATEGORIES
# =========================

def get_categories():

    print("📂 Получаем дерево категорий...")

    soup = get_soup(BASE)

    categories = []
    seen = set()

    # Основной блок категорий
    root = soup.select_one(
        "#categories_block_left"
    )

    if not root:

        print("⚠ Блок категорий не найден")

        return categories

    # Берём ВСЕ ссылки внутри дерева.
    # Поэтому автоматически получаем:
    # 1 уровень
    # 2 уровень
    # 3 уровень
    # и т.д.

    for a in root.select(
        "ul.tree a[href]"
    ):

        href = a.get("href", "").strip()

        if not href:
            continue

        href = absolute_url(href)

        # Только ссылки на категории этого сайта
        if "arizonesports.com.ua" not in href:
            continue

        # Не добавляем дубли
        if href in seen:
            continue

        seen.add(href)

        name = clean(
            a.get_text(" ", strip=True)
        )

        categories.append({
            "name": name,
            "url": href
        })

    print(
        f"📂 Найдено категорий: "
        f"{len(categories)}"
    )

    for category in categories:

        print(
            f"   • {category['name']}"
        )

    return categories


# =========================
# SHOW ALL
# =========================

def get_show_all_url(cat_url, soup):

    form = soup.select_one(
        "form.showall"
    )

    if not form:
        return None

    action = form.get("action", "").strip()

    if not action:
        action = cat_url

    action = absolute_url(action)

    params = {}

    # id_category
    id_category = form.select_one(
        "input[name='id_category']"
    )

    if id_category:

        value = id_category.get("value", "").strip()

        if value:
            params["id_category"] = value

    # количество товаров
    n_input = form.select_one(
        "input[name='n']"
    )

    if n_input:

        value = n_input.get("value", "").strip()

        if value:

            try:
                n = int(value)

                if n > 0:
                    params["n"] = str(n)

            except:
                pass

    if not params.get("n"):
        return None

    return action, params


# =========================
# PRODUCT LINKS
# =========================

def extract_product_links(soup):

    links = []

    seen = set()

    # Основной блок товаров PrestaShop
    selectors = [
        "ul.product_list li.ajax_block_product h5 a",
        "ul.product_list li.ajax_block_product a.product-name",
        "ul.product_list li.ajax_block_product a[href]",
        ".product-container a.product-name",
        ".product-container h5 a",
        ".product-container a[href]"
    ]

    for selector in selectors:

        for a in soup.select(selector):

            href = a.get("href", "").strip()

            if not href:
                continue

            href = absolute_url(href)

            if "arizonesports.com.ua" not in href:
                continue

            # Исключаем явно ненужные ссылки
            if "#" in href:
                href = href.split("#")[0]

            # Ссылка должна вести на товар
            # У сайта товарные URL заканчиваются .html
            if ".html" not in href:
                continue

            if href in seen:
                continue

            seen.add(href)
            links.append(href)

        if links:
            break

    return links


# =========================
# PARSE CATEGORY
# =========================

def parse_category(cat):

    cat_name = cat["name"]
    cat_url = cat["url"]

    print()
    print(
        f"📁 CATEGORY: {cat_name}"
    )

    first_page = get_soup(cat_url)

    if not first_page:

        print("⚠ Пустая страница")

        return []


    # =========================
    # ПЫТАЕМСЯ ПОЛУЧИТЬ ВСЕ
    # =========================

    show_all = get_show_all_url(
        cat_url,
        first_page
    )

    if show_all:

        action, params = show_all

        print(
            f"   📦 Пытаемся получить все товары: "
            f"n={params.get('n')}"
        )

        all_page = get_soup(
            action,
            params=params
        )

        links = extract_product_links(
            all_page
        )

        print(
            f"   🔗 Товаров найдено: {len(links)}"
        )

        if links:

            return links


    # =========================
    # FALLBACK — ПАГИНАЦИЯ
    # =========================

    print(
        "   ⚠ Показать все не сработало. "
        "Используем пагинацию."
    )

    links = []

    seen = set()

    # Первая страница
    page_links = extract_product_links(
        first_page
    )

    for link in page_links:

        if link not in seen:

            seen.add(link)
            links.append(link)


    # Ищем максимальную страницу
    max_page = 1

    for a in first_page.select(
        "#pagination_bottom a[href]"
    ):

        href = a.get("href", "")

        # Здесь PrestaShop использует
        # #/page-2
        m = re.search(
            r"#/page-(\d+)",
            href
        )

        if m:

            page = int(m.group(1))

            max_page = max(
                max_page,
                page
            )


    print(
        f"   📄 Страниц: {max_page}"
    )


    # Если есть страницы
    for page in range(
        2,
        max_page + 1
    ):

        # Важно:
        # hash #/page-2 сам сервер не получает.
        #
        # Поэтому пробуем стандартный
        # параметр page.
        url = cat_url

        soup = get_soup(
            url,
            params={
                "page": page
            }
        )

        page_links = extract_product_links(
            soup
        )

        print(
            f"   📄 Страница {page}: "
            f"{len(page_links)} товаров"
        )

        for link in page_links:

            if link not in seen:

                seen.add(link)
                links.append(link)


    print(
        f"   🔗 Всего товаров: {len(links)}"
    )

    return links


# =========================
# STATUS
# =========================

def parse_status(soup):

    status_el = soup.select_one(
        "#availability_value"
    )

    status_text = ""

    if status_el:

        status_text = clean(
            status_el.get_text()
        )


    # Если есть явный текст
    if status_text:

        return status_text


    # Проверяем schema.org
    availability = soup.select_one(
        "[itemprop='availability']"
    )

    if availability:

        href = (
            availability.get("href", "")
            .strip()
            .lower()
        )

        if "instock" in href:

            return "В наявності"

        if "outofstock" in href:

            return "Немає в наявності"


    # Если ничего нет
    return ""


# =========================
# COLORS
# =========================

def parse_colors(soup):

    colors = []

    seen = set()

    # Ищем поля атрибутов
    for fieldset in soup.select(
        "#attributes .attribute_fieldset"
    ):

        label = fieldset.select_one(
            ".attribute_label"
        )

        if not label:
            continue

        label_text = clean(
            label.get_text()
        ).lower()

        # Именно поле "Колір"
        if "колір" not in label_text:
            continue

        for li in fieldset.select(
            ".attribute_list li"
        ):

            text = clean(
                li.get_text(
                    " ",
                    strip=True
                )
            )

            if not text:
                continue

            if text in seen:
                continue

            seen.add(text)
            colors.append(text)


    return ", ".join(colors)


# =========================
# PRODUCT
# =========================

def parse_product(url):

    soup = get_soup(url)

    if not soup:

        return None


    # =========================
    # TITLE
    # =========================

    title_el = soup.select_one(
        "h1[itemprop='name']"
    )

    title = ""

    if title_el:

        title = clean(
            title_el.get_text()
        )


    # =========================
    # SKU
    # =========================

    sku_el = soup.select_one(
        "#product_reference "
        "span[itemprop='sku']"
    )

    sku = ""

    if sku_el:

        sku = clean(
            sku_el.get_text()
        )

    # fallback
    if not sku:

        sku_el = soup.select_one(
            "#product_reference .editable"
        )

        if sku_el:

            sku = clean(
                sku_el.get_text()
            )


    # =========================
    # STATUS
    # =========================

    status = parse_status(
        soup
    )


    # =========================
    # COLORS
    # =========================

    colors = parse_colors(
        soup
    )


    # =========================
    # RESULT
    # =========================

    return {
        "sku": sku,
        "title": title,
        "price_usd": "",
        "status": status,
        "colors": colors,
        "url": url
    }


# =========================
# MAIN
# =========================

def run_parser():

    if is_locked():

        print(
            "⚠ Парсер уже запущен."
        )

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
        # EXCEL
        # =========================

        wb = Workbook()

        ws = wb.active

        ws.title = "Products"

        ws.append([
            "SKU",
            "TITLE",
            "PRICE USD",
            "STATUS",
            "COLORS",
            "URL"
        ])


        # =========================
        # CATEGORIES
        # =========================

        categories = get_categories()


        if CATEGORY_LIMIT:

            categories = categories[
                :CATEGORY_LIMIT
            ]


        total_categories = len(
            categories
        )


        if total_categories == 0:

            print(
                "⚠ Категории не найдены."
            )

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return


        # =========================
        # COLLECT PRODUCT URLS
        # =========================

        all_product_urls = []

        seen_products = set()


        for i, category in enumerate(
            categories,
            1
        ):

            progress = int(
                i
                / total_categories
                * 40
            )


            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )


            product_urls = parse_category(
                category
            )


            for url in product_urls:

                if url in seen_products:
                    continue

                seen_products.add(url)

                all_product_urls.append(url)


            time.sleep(0.2)


        print()
        print(
            f"🛒 Всего уникальных товаров: "
            f"{len(all_product_urls)}"
        )


        # =========================
        # PRODUCT LIMIT
        # =========================

        if PRODUCT_LIMIT:

            all_product_urls = (
                all_product_urls[
                    :PRODUCT_LIMIT
                ]
            )

            print(
                f"⚙ PRODUCT_LIMIT: "
                f"{PRODUCT_LIMIT}"
            )


        total_products = len(
            all_product_urls
        )


        if total_products == 0:

            print(
                "⚠ Товары не найдены."
            )

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return


        # =========================
        # PARSE PRODUCTS
        # =========================

        seen_sku = set()


        for index, url in enumerate(
            all_product_urls,
            1
        ):

            progress = (
                40
                + int(
                    index
                    / total_products
                    * 60
                )
            )


            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )


            print(
                f"[{index}/{total_products}] "
                f"{url}"
            )


            try:

                item = parse_product(
                    url
                )

                if not item:

                    print(
                        "   ⚠ Не удалось "
                        "разобрать товар"
                    )

                    continue


                sku = item["sku"]


                # Если SKU есть —
                # не добавляем дубль
                if sku:

                    if sku in seen_sku:

                        print(
                            f"   ↪ Дубликат SKU: "
                            f"{sku}"
                        )

                        continue

                    seen_sku.add(sku)


                ws.append([
                    item["sku"],
                    item["title"],
                    item["price_usd"],
                    item["status"],
                    item["colors"],
                    item["url"]
                ])


                print(
                    f"   ✅ "
                    f"{item['sku']} | "
                    f"{item['status']} | "
                    f"{item['colors']}"
                )


            except Exception as e:

                print(
                    f"   ❌ Ошибка товара: "
                    f"{e}"
                )


            time.sleep(0.15)


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


        print()
        print(
            "✅ ГОТОВО ARIZONE_Sports"
        )

        print(
            f"📄 Файл: {FILE_PATH}"
        )

        print(
            f"🛒 Товаров: "
            f"{len(all_product_urls)}"
        )


    finally:

        set_lock(False)


# =========================
# START
# =========================

if __name__ == "__main__":

    run_parser()
