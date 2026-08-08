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

print("🔥 Харьковская 4421-4422 Jmax")

BASE = "https://www.jmaxtvshop.com.ua"

# =========================
# ⚙️ SWITCH
# =========================

# CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "4421-4422_Jmax_LIVE.xlsx"
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
    "User-Agent": "Mozilla/5.0",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
    "Connection": "keep-alive"
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

        age = time.time() - os.path.getmtime(
            LOCK_FILE
        )

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
# LOGIN
# =========================

def login():

    login_url = (
        BASE +
        "/index.php?route=account/login"
    )

    print("🔐 LOGIN...")

    for attempt in range(1, 4):

        try:

            r = session.get(
                login_url,
                timeout=30,
                allow_redirects=True
            )

            if r.status_code == 429:

                print(
                    f"⚠️ Jmax 429. "
                    f"Ждём 5 минут... "
                    f"({attempt}/3)"
                )

                if attempt < 3:

                    time.sleep(300)

                    continue

                return False

            if r.status_code != 200:

                print(
                    f"❌ LOGIN HTTP {r.status_code}"
                )

                return False

            soup = BeautifulSoup(
                r.text,
                "html.parser"
            )

            form = soup.select_one(
                'form[action*="route=account/login"]'
            )

            if not form:

                form = soup.select_one(
                    "form"
                )

            if not form:

                print(
                    "❌ LOGIN FORM NOT FOUND"
                )

                return False

            payload = {}

            for inp in form.select("input"):

                name = inp.get("name")

                if name:

                    payload[name] = inp.get(
                        "value",
                        ""
                    )

            payload["email"] = EMAIL
            payload["password"] = PASSWORD

            action = form.get(
                "action"
            )

            if not action:

                action = login_url

            if not action.startswith("http"):

                action = (
                    BASE +
                    "/" +
                    action.lstrip("/")
                )

            response = session.post(
                action,
                data=payload,
                headers={
                    "Referer": login_url
                },
                timeout=30,
                allow_redirects=True
            )

            if response.status_code == 429:

                print(
                    "⚠️ Jmax 429 после LOGIN. "
                    "Ждём 5 минут..."
                )

                if attempt < 3:

                    time.sleep(300)

                    continue

                return False

            account = session.get(
                BASE,
                timeout=30,
                allow_redirects=True
            )

            if account.status_code == 429:

                print(
                    "⚠️ Jmax 429 при проверке. "
                    "Ждём 5 минут..."
                )

                if attempt < 3:

                    time.sleep(300)

                    continue

                return False

            html = account.text.lower()

            if (
                "logout" in html
                or "account/logout" in html
                or "выйти" in html
                or "вихід" in html
            ):

                print("✅ LOGIN OK")

                return True

            print("⚠️ LOGIN CHECK")

            return True

        except Exception as e:

            print(
                f"❌ LOGIN ERROR: {e}"
            )

            if attempt < 3:

                time.sleep(300)

            else:

                return False

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

    for attempt in range(1, 4):

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
                    "⚠️ Jmax 429. "
                    "Ждём 5 минут..."
                )

                if attempt < 3:

                    time.sleep(300)

                    continue

        except:

            pass

        if attempt < 3:

            time.sleep(2)

    return BeautifulSoup(
        "",
        "html.parser"
    )


# =========================
# CLEAN
# =========================

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
# URL NORMALIZATION
# =========================

def normalize_url(url):

    if not url:
        return ""

    url = url.strip()

    if not url:
        return ""

    if url.startswith("#"):
        return ""

    if url.startswith("javascript"):
        return ""

    if not url.startswith("http"):

        url = (
            BASE +
            "/" +
            url.lstrip("/")
        )

    return url.split("#")[0]


# =========================
# PRODUCT URL CHECK
# =========================

def is_product_url(url):

    if not url:
        return False

    return (
        "route=product/product"
        in url
    )


# =========================
# CATEGORIES
# =========================

def get_categories():

    print("🌳 ПОИСК КАТЕГОРИЙ")

    soup = get_soup(BASE)

    if not soup or not soup.find_all(True):

        print(
            "❌ Главная страница не загрузилась"
        )

        return []

    categories = []
    seen = set()

    def add_category(url):

        url = normalize_url(url)

        if not url:
            return

        if (
            "route=product/category"
            not in url
        ):
            return

        url = re.sub(
            r"[&?]page=\d+",
            "",
            url
        )

        if url in seen:
            return

        seen.add(url)

        categories.append(url)

    for tag in soup.find_all(True):

        add_category(
            tag.get("href")
        )

        add_category(
            tag.get("data-href")
        )

    print(
        f"📂 Categories: {len(categories)}"
    )

    return categories


# =========================
# LAST PAGE
# =========================

def get_last_page(soup):

    pages = [1]

    for a in soup.select(
        "ul.pagination a"
    ):

        href = a.get(
            "href",
            ""
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
# PRODUCT LINKS FROM PAGE
# =========================

def get_product_links(soup):

    products = []
    seen = set()

    if not soup:
        return products

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = normalize_url(
            a.get("href")
        )

        if not is_product_url(href):
            continue

        if href in seen:
            continue

        seen.add(href)

        products.append(href)

    return products


# =========================
# PARSE CATEGORY
# =========================

def parse_category(cat_url):

    all_products = []
    seen = set()

    first_page = get_soup(
        cat_url
    )

    if not first_page or not first_page.find_all(True):

        return all_products

    last_page = get_last_page(
        first_page
    )

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:

            url = cat_url

        else:

            separator = (
                "&"
                if "?" in cat_url
                else "?"
            )

            url = (
                cat_url +
                f"{separator}page={page}"
            )

        soup = get_soup(
            url
        )

        if not soup or not soup.find_all(True):

            break

        products = get_product_links(
            soup
        )

        if not products:

            break

        for product_url in products:

            if product_url in seen:
                continue

            seen.add(product_url)

            all_products.append(
                product_url
            )

    return all_products


# =========================
# SITEMAP
# =========================

def get_sitemap_urls():

    print("🗺 ПРОВЕРЯЕМ SITEMAP")

    sitemap_candidates = [

        BASE + "/sitemap.xml",

        BASE + "/sitemap_index.xml",

        BASE + "/index.php?route=feed/google_sitemap",

        BASE + "/index.php?route=extension/feed/google_sitemap"
    ]

    urls = []
    seen = set()

    for sitemap_url in sitemap_candidates:

        soup = get_soup(
            sitemap_url
        )

        if not soup:
            continue

        links = soup.find_all(
            "loc"
        )

        for loc in links:

            url = clean(
                loc.get_text()
            )

            url = normalize_url(
                url
            )

            if not url:
                continue

            if not is_product_url(url):
                continue

            if url in seen:
                continue

            seen.add(url)

            urls.append(url)

    print(
        f"🗺 Sitemap products: {len(urls)}"
    )

    return urls


# =========================
# SEARCH ENGINE
# =========================

def get_search_product_links():

    print("🔎 ПРОВЕРЯЕМ ПОИСК")

    urls = []

    # Основной поиск OpenCart.
    # Сам по себе он не знает все товары,
    # поэтому используем его как дополнительный источник.

    search_urls = [
        BASE +
        "/index.php?route=product/search&search=",

        BASE +
        "/index.php?route=product/search&filter_name="
    ]

    # Сюда попадут только ссылки,
    # которые реально были найдены поиском.
    #
    # Пока не перебираем артикулы,
    # потому что это может создать
    # огромное количество запросов.

    for search_base in search_urls:

        soup = get_soup(
            search_base
        )

        if not soup:
            continue

        products = get_product_links(
            soup
        )

        for url in products:

            if url not in urls:

                urls.append(url)

    print(
        f"🔎 Search products: {len(urls)}"
    )

    return urls

# =========================
# SEARCH PRODUCTS BY SKU
# =========================

def search_products_by_sku(
    start_sku=1,
    end_sku=10000,
    delay=0.3
):

    print(
        f"🔎 ПОИСК ТОВАРОВ ПО АРТИКУЛАМ "
        f"{start_sku}-{end_sku}"
    )

    found_urls = []
    found_seen = set()

    checked = 0

    for sku in range(
        start_sku,
        end_sku + 1
    ):

        checked += 1

        search_url = (
            BASE +
            "/index.php?route=product/search&search="
            + str(sku)
        )

        soup = get_soup(
            search_url
        )

        if not soup:
            time.sleep(delay)
            continue

        products = get_product_links(
            soup
        )

        for product_url in products:

            if product_url in found_seen:
                continue

            found_seen.add(
                product_url
            )

            found_urls.append(
                product_url
            )

            print(
                f"➕ Найден SKU {sku}: "
                f"{product_url}"
            )

        if checked % 100 == 0:

            print(
                f"🔎 Проверено артикулов: "
                f"{checked}/{end_sku - start_sku + 1} "
                f"| Найдено товаров: "
                f"{len(found_urls)}"
            )

        time.sleep(delay)

    print(
        f"✅ ПОИСК ЗАВЕРШЁН. "
        f"Найдено дополнительных товаров: "
        f"{len(found_urls)}"
    )

    return found_urls


# =========================
# PRODUCT
# =========================

def parse_product(url):

    soup = get_soup(
        url
    )

    if not soup:

        return [
            "",
            "",
            "",
            "",
            url
        ]

    h1 = soup.select_one(
        "h1"
    )

    if not h1:

        return [
            "",
            "",
            "",
            "",
            url
        ]

    title = clean(
        h1.get_text()
    )

    # =========================
    # SKU
    # =========================

    sku = ""

    selectors = [

        ".product-data__item.model",

        ".product-data__item",

        ".model",

        ".product-model",

        ".product-info .model",

        ".product-page__model",

        "[class*='model']",

        "[class*='sku']",

        "[class*='article']",

        "[class*='articul']"
    ]

    for selector in selectors:

        tags = soup.select(
            selector
        )

        for tag in tags:

            text = clean(
                tag.get_text(
                    " ",
                    strip=True
                )
            )

            if not text:
                continue

            test = text

            test = re.sub(
                r"^(Код товара|Артикул|Код|SKU|Модель|Model)\s*[:\-]?\s*",
                "",
                test,
                flags=re.I
            ).strip()

            m = re.search(
                r"\b(\d{3,10})\b",
                test
            )

            if m:

                sku = m.group(1)

                break

        if sku:
            break

    # =========================
    # PAGE TEXT
    # =========================

    if not sku:

        page_text = clean(
            soup.get_text(
                " ",
                strip=True
            )
        )

        patterns = [

            r"(?:Код\s*товара)\s*[:\-]?\s*(\d{3,10})",

            r"(?:Артикул)\s*[:\-]?\s*(\d{3,10})",

            r"(?:Код)\s*[:\-]?\s*(\d{3,10})",

            r"(?:SKU)\s*[:\-]?\s*(\d{3,10})",

            r"(?:Модель)\s*[:\-]?\s*(\d{3,10})",

            r"(?:Model)\s*[:\-]?\s*(\d{3,10})"
        ]

        for pattern in patterns:

            m = re.search(
                pattern,
                page_text,
                flags=re.I
            )

            if m:

                sku = m.group(1)

                break

    # =========================
    # JSON-LD
    # =========================

    if not sku:

        for script in soup.select(
            'script[type="application/ld+json"]'
        ):

            text = script.get_text(
                strip=True
            )

            if not text:
                continue

            m = re.search(
                r'"sku"\s*:\s*"([^"]+)"',
                text,
                flags=re.I
            )

            if m:

                sku = clean(
                    m.group(1)
                )

                break

    # =========================
    # PRICE
    # =========================

    price = ""

    price_tag = soup.select_one(
        ".product-page__price"
    )

    if price_tag:

        price = clean(
            price_tag.get_text()
        )

    # =========================
    # STATUS
    # =========================

    status = ""

    btn = soup.select_one(
        "#button-cart span"
    )

    if btn:

        status = clean(
            btn.get_text()
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
                "❌ Не удалось авторизоваться"
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
            "PRICE",
            "STATUS",
            "URL"
        ])

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        # =========================
        # COLLECT PRODUCTS
        # =========================

        product_urls = []
        product_seen = set()

        # =========================
        # 1. CATEGORIES
        # =========================

        cats = get_categories()

        if CATEGORY_LIMIT:

            cats = cats[
                :CATEGORY_LIMIT
            ]

        total_categories = len(cats)

        for i, cat in enumerate(
            cats,
            1
        ):

            links = parse_category(
                cat
            )

            for url in links:

                if url in product_seen:
                    continue

                product_seen.add(url)

                product_urls.append(
                    url
                )

        print(
            f"📦 Товаров из категорий: "
            f"{len(product_urls)}"
        )

        # =========================
        # 2. SITEMAP
        # =========================

        sitemap_products = (
            get_sitemap_urls()
        )

        for url in sitemap_products:

            if url in product_seen:
                continue

            product_seen.add(url)

            product_urls.append(
                url
            )

        print(
            f"📦 Всего ссылок на товары: "
            f"{len(product_urls)}"
        )

        # =========================
        # 3. ПОИСК ТОВАРОВ ПО АРТИКУЛАМ
        # =========================
        
        search_products = search_products_by_sku(
            start_sku=1,
            end_sku=3000,
            delay=0.3
        )
        
        additional_products = 0
        
        for url in search_products:
        
            if url in product_seen:
                continue
        
            product_seen.add(
                url
            )
        
            product_urls.append(
                url
            )
        
            additional_products += 1
        
        print(
            f"➕ Дополнительных товаров: "
            f"{additional_products}"
        )
        
        print(
            f"📦 ИТОГО ТОВАРОВ К ПАРСИНГУ: "
            f"{len(product_urls)}"
        )

        
        # =========================
        # PARSE PRODUCTS
        # =========================

        total_products = len(
            product_urls
        )

        for i, product_url in enumerate(
            product_urls,
            1
        ):

            progress = (
                int(
                    i /
                    total_products *
                    100
                )
                if total_products
                else 100
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )

            item = parse_product(
                product_url
            )

            if not item:
                continue

            sku, title, price, status, url = item

            if not title:
                continue

            ws.append([
                sku,
                title,
                price,
                status,
                url
            ])

        # =========================
        # SAVE
        # =========================

        tmp = FILE_PATH + ".tmp"

        wb.save(
            tmp
        )

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
            "Харьковская 4421-4422 Jmax"
        )

    finally:

        set_lock(False)


if __name__ == "__main__":

    run_parser()
