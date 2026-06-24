import os
import json
import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook


# =========================
# CONFIG
# =========================

BASE_DIR = os.path.abspath("output/4425-4426_Gold_Top")

FILE_PATH = os.path.join(BASE_DIR, "Харьковская_4425-4426_Gold_Top_LIVE.xlsx")
STATUS_PATH = os.path.join(BASE_DIR, "status.json")
LOCK_FILE = os.path.join(BASE_DIR, "lock.txt")

CATEGORY_LIMIT = None  # или 1 для теста
#CATEGORY_LIMIT = 1  # или 1 для теста

LOGIN_URL = "https://www.gold-tor.com.ua/index.php?route=account/login"

EMAIL = "Sawrun_05@icloud.com"
PASSWORD = "18022021"

BASE_URL = "https://www.gold-tor.com.ua"

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
# STATUS
# =========================

def save_status(data):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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
        self.wb.save(self.path)


# =========================
# PARSING HELPERS
# =========================

def get_soup(url):
    try:
        r = session.get(url, headers=HEADERS, timeout=20)
        r.raise_for_status()
        return BeautifulSoup(r.text, "html.parser")
    except Exception as e:
        print("⚠️ GET ERROR:", url, e)
        return BeautifulSoup("", "html.parser")


def parse_product(url):
    soup = get_soup(url)

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

    for a in soup.select("div.product-name a[href]"):
        href = a.get("href")
    
        if href:
            links.add(href.split("?")[0])


    print("FOUND LINKS:", len(links))
    return list(links)


def get_pages(cat_url):
    pages = []
    page = 1

    for _ in range(50):  # защита от бесконечного цикла
        url = f"{cat_url}?page={page}"
        soup = get_soup(url)

        items = soup.select("div.product-item")

        if not items:
            break

        pages.append(url)
        page += 1

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
    print("🚀 STARTED MAIN")
    create_lock()

    if os.path.exists(FILE_PATH):
        os.remove(FILE_PATH)

    try:
        os.makedirs(BASE_DIR, exist_ok=True)

        if not login():
            return

        excel = ExcelWriter(FILE_PATH)
        seen = set()
        all_links = set()
        found = 0
        written = 0

        categories = get_categories()

        if CATEGORY_LIMIT:
            categories = categories[:CATEGORY_LIMIT]

        print(f"📦 Categories: {len(categories)}")

        for cat_name, cat_url in categories:
            print(f"\n📁 CATEGORY: {cat_name}")

            pages = [cat_url]

            more_pages = get_pages(cat_url)
            if more_pages:
                pages.extend(more_pages)

            for page in pages:
                print(f"➡ PAGE: {page}")


                product_links = get_products_from_category(page)

                found += len(product_links)

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
                        
                        written += 1
                        
                        print("✔", data["name"])

                    except Exception as e:
                        print("❌ product error:", e)

                excel.save()

            save_status({
                "last_category": cat_name,
                "time": time.time()
            })

        print("FOUND:", found)
        print("WRITTEN:", written)
        print("\n✅ DONE")

    finally:
        remove_lock()


if __name__ == "__main__":
    main()
