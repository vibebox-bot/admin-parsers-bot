import os
import re
import json
import time
import random
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook, load_workbook
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from datetime import datetime

visited_categories = set()
TOTAL_PRODUCTS = 0
seen = set()

print("🟢 228_Rainberg")

# =====================================================
# НАСТРОЙКИ
# =====================================================

BASE_URL = "https://rainberg.org.ua"
CATALOG_URL = "https://rainberg.org.ua/ua/product_list"

OUTPUT_DIR = os.path.abspath("output/228_Rainberg")
FILE_PATH = os.path.join(OUTPUT_DIR, "228_Rainberg_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

#CATEGORY_LIMIT = None
CATEGORY_LIMIT = 1

import sys
USER = sys.argv[1] if len(sys.argv) > 1 else ""

os.makedirs(OUTPUT_DIR, exist_ok=True)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/137.0 Safari/537.36"
    )
}

# =====================================================
# SESSION
# =====================================================

retry = Retry(
    total=5,
    connect=5,
    read=5,
    backoff_factor=1,
    status_forcelist=[429, 500, 502, 503, 504],
)

adapter = HTTPAdapter(max_retries=retry)

session = requests.Session()
session.headers.update(HEADERS)
session.mount("https://", adapter)
session.mount("http://", adapter)

# =====================================================
# STATUS
# =====================================================

def set_status(running=True, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path
    }

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        # lock старше часа = считаем зависшим
        if age > 3600:
            os.remove(LOCK_FILE)
            return False

        return True

    except:
        return False


def set_lock(state):

    if state:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)

# =====================================================
# EXCEL
# =====================================================

def get_workbook():

    if os.path.exists(FILE_PATH):
        os.remove(FILE_PATH)

    wb = Workbook()
    ws = wb.active

    ws.append([
        "Название",
        "Артикул",
        "Цена",
        "Наличие",
        "Ссылка"
    ])

    return wb, ws

# =====================================================
# REQUEST
# =====================================================

def get_soup(url):

    attempt = 0

    while attempt < 5:
        attempt += 1

        try:

            r = session.get(url, timeout=15)

            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

            print(f"⚠️ HTTP {r.status_code}")

        except Exception as e:
            print("❌", e)

        time.sleep(random.uniform(2, 4))

    return None

# =====================================================
# КАТЕГОРИИ
# =====================================================

def get_categories():
    soup = get_soup(CATALOG_URL)
    if soup is None:
        print("❌ Не удалось загрузить каталог")
        return []

    categories = []
    blocks = soup.select("ul.b-product-groups-gallery li")

    for block in blocks:
        a = block.select_one("a.b-product-groups-gallery__title")
        if not a:
            continue

        href = a.get("href")
        if not href:
            continue

        categories.append(urljoin(BASE_URL, href))

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

    print(f"📂 Категорий найдено: {len(categories)}")
    for i, c in enumerate(categories):
        #print(f"{i+1}. {c}")

    return categories


def get_subcategories(category_url):
    soup = get_soup(category_url)
    subs = []

    blocks = soup.select("ul.b-product-groups-gallery li")

    for block in blocks:
        a = block.select_one("a.b-product-groups-gallery__title")
        if not a:
            continue

        href = a.get("href")
        if href:
            subs.append(urljoin(BASE_URL, href))

    return subs


def process_category(category_url, ws, wb):

    if category_url in visited_categories:
        return

    visited_categories.add(category_url)

    #print(f"📂 CATEGORY: {category_url}")

    # ==========================================
    # Сначала товары текущей категории
    # ==========================================

    pages = get_pages(category_url)

    for page in pages:

        #print("\n📄 PAGE:", page)

        products = get_products(page)

        for product_url in products:

            try:
                product = parse_product(product_url)

                key = product["sku"].strip() if product["sku"] else product["url"]
                
                if key in seen:
                    continue
                
                seen.add(key)
                
                save_product(ws, product)

                time.sleep(random.uniform(0.5, 1.2))

            except Exception as e:
                print("Ошибка товара:", product_url, e)

        time.sleep(random.uniform(1, 2))

    # ==========================================
    # Потом все вложенные категории
    # ==========================================

    subcategories = get_subcategories(category_url)

    if subcategories:

        #print(f"📁 Подкатегорий: {len(subcategories)}")

        for sub in subcategories:
            process_category(sub, ws, wb)

# =====================================================
# ПАГИНАТОР
# =====================================================

def get_pages(category_url):

    soup = get_soup(category_url)

    pages = [category_url]

    pager = soup.select_one("[data-pagination-pages-count]")

    if not pager:
        #print("Страниц: 1")
        return pages

    try:
        last_page = int(pager["data-pagination-pages-count"])
    except:
        last_page = 1

    for i in range(2, last_page + 1):
        pages.append(category_url.rstrip("/") + f"/page_{i}")

    #print(f"Страниц: {last_page}")

    return pages

# =====================================================
# ТОВАРЫ
# =====================================================

def get_products(page_url):
    soup = get_soup(page_url)
    products = []

    cards = soup.select("li[data-product-id]")

    for card in cards:
        a = card.select_one("a.b-product-gallery__title")
        if not a:
            continue

        href = a.get("href")
        if href:
            products.append(urljoin(BASE_URL, href))

    #print(f"🛒 Товаров на странице: {len(products)}")

    if len(products) == 0:
        #print("⚠️ НЕТ ТОВАРОВ — возможно сайт блокирует или сломался селектор")

    return products

# =====================================================
# ПАРС ТОВАРА
# =====================================================

def clean_price(text):
    if not text:
        return ""
    text = text.replace("\xa0", "")
    text = text.replace("₴", "")
    text = text.replace("/шт.", "")
    return text.strip()


def get_text(el):
    if not el:
        return ""
    return el.get_text(" ", strip=True)


def parse_product(url):
    soup = get_soup(url)

    title = get_text(soup.select_one('span[data-qaid="product_name"]'))
    sku = get_text(soup.select_one('span[data-qaid="product_code"]'))

    price = ""
    price_tag = soup.select_one('span[data-qaid="wholesale_price"]')
    if price_tag:
        price = clean_price(price_tag.get_text(strip=True))

    availability = ""
    state = soup.select_one('[data-qaid="presence_data"]')
    if state:
        availability = state.get_text(" ", strip=True)

    return {
        "title": title,
        "sku": sku,
        "price": price,
        "availability": availability,
        "url": url
    }

# =====================================================
# EXCEL SAVE
# =====================================================

def save_product(ws, product):

    global TOTAL_PRODUCTS

    ws.append([
        product["title"],
        product["sku"],
        product["price"],
        product["availability"],
        product["url"]
    ])

    TOTAL_PRODUCTS += 1

    if TOTAL_PRODUCTS % 50 == 0:
        #print(f"📦 Сохранено товаров: {TOTAL_PRODUCTS}")
# =====================================================
# MAIN
# =====================================================

def main():
    #print("🔥 ENTER MAIN()")
    #print("🚀 PARSER STARTED")
    #print(f"👤 USER: {USER}")

    if is_locked():
        #print("⛔ Уже запущен")
        return

    set_lock(True)

    try:

        set_status(
            running=True,
            user=USER,
            file_path=FILE_PATH
        )

        wb, ws = get_workbook()

        categories = get_categories()

        for category_index, category_url in enumerate(categories, 1):

            #print("\n" + "=" * 70)
            #print(f"Категория {category_index}/{len(categories)}")
            #print(category_url)

            process_category(category_url, ws, wb)

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        wb.save(FILE_PATH)

        set_status(
            running=False,
            user=USER,
            file_path=FILE_PATH
        )

        print(f"✅ Готово. Всего товаров: {saved_count}")


    finally:
        set_lock(False)
    
# =====================================================
# START
# =====================================================

def run_parser():
    main()


if __name__ == "__main__":
    run_parser()

