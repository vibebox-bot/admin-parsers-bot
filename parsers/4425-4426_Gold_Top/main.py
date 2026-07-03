import os
import sys
import json
import time
import requests

from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import datetime

import pytz

USER = sys.argv[1] if len(sys.argv) > 1 else "unknown"

OUTPUT_DIR = os.path.abspath("output/4425-4426_Gold_Top")
FILE_PATH = os.path.join(OUTPUT_DIR, "4425-4426_Gold_Top_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "../../")))

os.makedirs(OUTPUT_DIR, exist_ok=True)

#CATEGORY_LIMIT = None  # или 1 для теста
CATEGORY_LIMIT = 1  # или 1 для теста

LOGIN_URL = "https://www.gold-tor.com.ua/index.php?route=account/login"

EMAIL = "Sawrun_05@icloud.com"
PASSWORD = "18022021"

BASE_URL = "https://www.gold-tor.com.ua"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()


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


# =========================
# LOCK
# =========================

def create_lock():
    if os.path.exists(LOCK_FILE):
        print("⚠️ OLD LOCK FOUND → removing")
        os.remove(LOCK_FILE)

    with open(LOCK_FILE, "w") as f:
        f.write("1")


def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


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
# EXCEL WRITER
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
        tmp_path = self.path + ".tmp"
        self.wb.save(tmp_path)
        os.replace(tmp_path, self.path)

# =========================
# PARSING HELPERS
# =========================

def get_soup(url):
    try:
        r = session.get(url, headers=HEADERS, timeout=20)

        if not r or not r.text:
            return None

        if r.status_code != 200:
            print(f"❌ HTTP {r.status_code}: {url}")
            return None

        if len(r.text) < 500:
            print(f"⚠️ EMPTY PAGE: {url}")
            return None

        return BeautifulSoup(r.text, "html.parser")

    except Exception as e:
        print(f"⚠️ REQUEST FAIL: {url} -> {e}")
        return None

def parse_product(url):
    soup = get_soup(url)

    if soup is None:
        return {
            "name": "",
            "price": "",
            "article": "",
            "stock": "",
            "url": url
        }

    # NAME
    name_tag = soup.select_one("h1.h2")
    name = name_tag.get_text(strip=True) if name_tag else ""

    # ARTICLE
    article_tag = soup.select_one("span.text-danger")
    article = article_tag.get_text(strip=True) if article_tag else ""

    # PRICE
    price_tag = soup.select_one(".h2.m-0.text-nowrap")
    price = price_tag.get_text(strip=True) if price_tag else ""

    # STOCK (как на сайте)
    stock_tag = soup.select_one(".alert")
    stock = stock_tag.get_text(" ", strip=True) if stock_tag else ""

    return {
        "name": name,
        "price": price,
        "article": article,
        "stock": stock,
        "url": url
    }

def get_products_from_category(url):
    soup = get_soup(url)

    links = set()

    for item in soup.select("div.product-layout, div.product-item, div.product-thumb"):
        a = item.select_one("div.product-image a[href]")
        if a:
            links.add(a["href"].split("?")[0])

    return list(links)


def get_pages(cat_url):
    pages = []
    seen = set()

    # всегда добавляем limit=100
    base = cat_url
    if "limit=" not in base:
        base += "&limit=100" if "?" in base else "?limit=100"

    soup = get_soup(base)
    if not soup:
        return [base]

    pages.append(base)
    seen.add(base)

    # собираем ВСЕ реальные страницы из пагинации
    for a in soup.select(".pagination a, ul.pagination a"):
        href = a.get("href")
        if not href:
            continue

        # нормализация ссылки
        page_url = href.split("&limit")[0]

        if page_url not in seen:
            seen.add(page_url)
            pages.append(page_url)

    return pages
    
# =========================
# CATEGORY LIST
# =========================

def get_categories():
    url = BASE_URL + "/nstrumenti"
    soup = get_soup(url)

    cats = soup.select("#d_category_menu_list a")

    result = []

    for c in cats:
        href = c.get("href")
        name = c.text.strip()

        if href:
            result.append((name, href))

    return result


# =========================
# MAIN
# =========================
def main():

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    set_status(
        running=True,
        user=USER,
        file_path=FILE_PATH
    )
        
    create_lock()    

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    excel = ExcelWriter(FILE_PATH)

    try:
        if not login():
            print("❌ LOGIN FAILED → STOP")
            return

        seen = set()

        categories = get_categories()
        done = 0
        total = len(categories)

        if CATEGORY_LIMIT:
            categories = categories[:CATEGORY_LIMIT]

        for i, (cat_name, cat_url) in enumerate(categories, start=1):
            print(f"\n📁 CATEGORY: {cat_name}")
        
            pages = get_pages(cat_url)
        
            for page in pages:
        
                product_links = get_products_from_category(page)
                #found += len(product_links)
        
                for link in product_links:
                    try:
                        data = parse_product(link)
        
                        url = data["url"].split("?")[0]
        
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
        
                        #written += 1   
        
                    except Exception as e:
                        print("❌ PRODUCT ERROR:", e)
        
                excel.save()


        set_status(
            running=False,
            user=USER,
            file_path=FILE_PATH
        )
        
        
        print("✅ DONE")

    except Exception as e:
        print("🔥 FATAL ERROR:", e)

        set_status(
            running=False,
            user=USER,
            file_path=FILE_PATH
        )
         

    finally:
        excel.save()
        remove_lock()


if __name__ == "__main__":
    main()
