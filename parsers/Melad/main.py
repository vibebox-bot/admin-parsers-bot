import os
import json
import time
import requests
import random
from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys
from datetime import datetime

USER = sys.argv[1] if len(sys.argv) > 1 else "-"


# =========================
# CONFIG
# =========================

BASE_DIR = os.path.abspath("output/Melad")

FILE_PATH = os.path.join(BASE_DIR, "Melad_LIVE.xlsx")
STATUS_PATH = os.path.join(BASE_DIR, "status.json")
LOCK_FILE = os.path.join(BASE_DIR, "lock.txt")

CATEGORY_LIMIT = None  # или 1 для теста
#CATEGORY_LIMIT = 1  # или 1 для теста

LOGIN_URL = "https://melad.com.ua/login/"
BASE_URL = "https://melad.com.ua"

EMAIL = "titorangelina@gmail.com"
PASSWORD = "18022021"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()


def save_status(running=False, progress=0, user="", file_path=""):
    os.makedirs(BASE_DIR, exist_ok=True)

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
# LOCK
# =========================

def create_lock():
    if os.path.exists(LOCK_FILE):
        #print("❌ Already running")
        exit()

    with open(LOCK_FILE, "w") as f:
        f.write("locked")


def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# HTTP
# =========================

def get_soup(url, retries=3):
    for i in range(retries):
        try:
            r = session.get(url, headers=HEADERS, timeout=20)
            r.raise_for_status()
            return BeautifulSoup(r.text, "html.parser")

        except Exception as e:
            #print(f"⚠️ Retry {i+1}/{retries} for {url}: {e}")
            time.sleep(2)

    return BeautifulSoup("", "html.parser")


# =========================
# LOGIN
# =========================

def login():
    #print("🔐 LOGIN...")

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
# CATEGORIES
# =========================
def get_categories():

    soup = get_soup(BASE_URL)

    categories = []
    seen = set()

    def walk(li):

        # ссылка текущей категории
        a = li.select_one(":scope > a[href]")

        if a:

            href = a.get("href", "").strip()

            if href \
                and "javascript" not in href \
                and href not in seen:

                if href.startswith("/"):
                    href = BASE_URL + href

                seen.add(href)
                categories.append((a.get_text(" ", strip=True), href))

        # ищем все вложенные li (любая глубина)
        for child in li.select(":scope > .dropdown-menu li"):
            walk(child)

    for li in soup.select("#menu > ul > li.has-children"):
        walk(li)

    # одиночные категории без подкатегорий
    for a in soup.select("#menu > ul > li:not(.has-children) > a[href]"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if "javascript" in href:
            continue

        if href.startswith("/"):
            href = BASE_URL + href

        if href not in seen:
            seen.add(href)
            categories.append((a.get_text(" ", strip=True), href))

    print("📂 TOTAL CATEGORIES:", len(categories))

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
    #print(f"📦 PARSING CATEGORY: {url}")

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

    #print(f"   → FOUND PRODUCTS: {len(products)}")

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
        tmp = self.path + ".tmp"
        self.wb.save(tmp)
        os.replace(tmp, self.path)
        

# =========================
# MAIN
# =========================

def main():

    create_lock()
    
    os.makedirs(BASE_DIR, exist_ok=True)
    
    save_status(
        running=True,
        progress=0,
        user=USER,
        file_path=FILE_PATH
    )
    
    print("🚀 STARTED MELAD")

    if not login():
        remove_lock()
    
        save_status(
            running=False,
            progress=0,
            user=USER,
            file_path=FILE_PATH
        )
    
        return
    
    excel = ExcelWriter(FILE_PATH)
    #print("EXCEL FILE:", FILE_PATH)
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

            time.sleep(random.uniform(0.5, 1.5))
            
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

                progress = int((written / max(found, 1)) * 100)
        
                save_status(
                    running=True,
                    progress=progress,
                    user=USER,
                    file_path=FILE_PATH
                )

        

        excel.save()

    save_status(
        running=False,
        progress=100,
        user=USER,
        file_path=FILE_PATH
    )
    
    print("✅ DONE")

    remove_lock()


if __name__ == "__main__":
    main()
