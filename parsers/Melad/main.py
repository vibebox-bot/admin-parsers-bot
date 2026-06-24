import os
import json
import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook


# =========================
# CONFIG
# =========================

BASE_DIR = os.path.abspath("output/Melad")

FILE_PATH = os.path.join(BASE_DIR, "Melad_live.xlsx")
STATUS_PATH = os.path.join(BASE_DIR, "status.json")
LOCK_FILE = os.path.join(BASE_DIR, "lock.txt")

#CATEGORY_LIMIT = None  # или 1 для теста
CATEGORY_LIMIT = 1  # или 1 для теста

LOGIN_URL = "https://melad.com.ua/login/"
BASE_URL = "https://melad.com.ua"

EMAIL = "titorangelina@gmail.com"
PASSWORD = "18022021"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()


# =========================
# LOCK
# =========================

def create_lock():
    if os.path.exists(LOCK_FILE):
        print("❌ Already running")
        exit()

    with open(LOCK_FILE, "w") as f:
        f.write("locked")


def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# HTTP
# =========================

def get_soup(url):
    r = session.get(url, headers=HEADERS)
    return BeautifulSoup(r.text, "html.parser")


# =========================
# LOGIN
# =========================

def login():
    print("🔐 LOGIN...")

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    r = session.post(LOGIN_URL, data=payload, headers=HEADERS)

    if "logout" in r.text.lower() or "account" in r.text.lower():
        print("✅ LOGIN SUCCESS")
        return True

    print("❌ LOGIN FAILED")
    return False


# =========================
# CATEGORIES (ТОЛЬКО ГЛАВНЫЕ)
# =========================

def get_categories():
    print("📂 GET CATEGORIES...")

    soup = get_soup(BASE_URL)

    categories = []

    # ищем главное меню
    menu_items = soup.select("li.has-children > a")

    for item in menu_items:
        name = item.get_text(strip=True)

        # чистим лишние иконки/стрелки
        name = name.replace("›", "").replace("›", "").strip()

        url = item.get("href")

        if not url:
            continue

        # приводим относительные ссылки к абсолютным
        if url.startswith("/"):
            url = BASE_URL + url

        # фильтр от мусора (если вдруг попадётся вложенное)
        if "javascript" in url:
            continue

        categories.append((name, url))

    # убираем дубликаты (очень важно для таких меню)
    categories = list(dict.fromkeys(categories))

    print(f"✅ FOUND CATEGORIES: {len(categories)}")

    return categories


# =========================
# PAGES (ПАГИНАЦИЯ)
# =========================

def get_pages(url):
    soup = get_soup(url)

    pages = set()
    pages.add(url)

    for a in soup.select("div.pagination_wrap a[href]"):
        href = a.get("href")

        if href and "page=" in href:
            pages.add(href)

    return sorted(list(pages))

# =========================
# PRODUCTS FROM CATEGORY
# =========================

def get_products_from_category(url):
    print(f"📦 PARSING CATEGORY: {url}")

    soup = get_soup(url)

    products = []

    items = soup.select("div.product-layout.product-grid")

    for item in items:

        a_tag = item.select_one("div.image a")
        if not a_tag:
            continue

        link = a_tag.get("href")

        name_tag = item.select_one("div.caption a")
        name = name_tag.get_text(strip=True) if name_tag else ""

        price_tag = item.select_one("p.price")
        price = price_tag.get_text(strip=True) if price_tag else ""

        sku_tag = item.select_one("span.kod_sku b")
        sku = sku_tag.get_text(strip=True) if sku_tag else ""

        btn = item.select_one("button.add_to_cart")
        stock = btn.get_text(strip=True) if btn else ""

        products.append({
            "url": link,
            "name": name,
            "price": price,
            "article": sku,
            "stock": stock
        })

    print(f"   → FOUND PRODUCTS: {len(products)}")

    return products


# =========================
# PRODUCT PARSER
# =========================

def parse_product(url):
    # TODO: вставим позже
    return {
        "name": "",
        "price": "",
        "article": "",
        "stock": "",
        "url": url
    }


# =========================
# EXCEL
# =========================

class ExcelWriter:
    def __init__(self, path):
        self.path = path
        self.wb = Workbook()
        self.ws = self.wb.active

        self.ws.append([
            "Name",
            "Price",
            "Article",
            "Stock",
            "URL"
        ])

    def add(self, name, price, article, stock, url):
        self.ws.append([name, price, article, stock, url])

    def save(self):
        self.wb.save(self.path)


# =========================
# MAIN
# =========================

def main():
    create_lock()

    os.makedirs(BASE_DIR, exist_ok=True)

    print("🚀 STARTED MELAD")

    if not login():
        return

    excel = ExcelWriter(FILE_PATH)
    print("EXCEL FILE:", FILE_PATH)
    excel.save()

    seen = set()
    found = 0
    written = 0

    categories = get_categories()
    
    if CATEGORY_LIMIT is not None:
        categories = categories[:CATEGORY_LIMIT]

    for cat_name, cat_url in categories:

        pages = get_pages(cat_url)

        for page in pages:

            product_links = get_products_from_category(page)
            found += len(product_links)

            for data in product_links:


                url = data["url"].split("?")[0].rstrip("/").lower()
                
                if url in seen:
                    continue
                
                seen.add(url)
                
            
                excel.add(
                    data["name"],
                    data["price"],
                    data["article"],
                    data["stock"],
                    data["url"]
                )
            
                written += 1

        excel.save()

    print("FOUND:", found)
    print("WRITTEN:", written)
    print("✅ DONE")

    remove_lock()


if __name__ == "__main__":
    main()
