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
    r = session.get(url, headers=HEADERS)
    return BeautifulSoup(r.text, "html.parser")


def parse_product(url):
    soup = get_soup(url)

    # NAME
    name = soup.select_one("h1")
    name = name.get_text(strip=True) if name else ""

    # PRICE
    price = soup.select_one(".h2")
    price = price.get_text(strip=True) if price else ""

    # ARTICLE / CODE
    article = soup.select_one(".text-danger")
    article = article.get_text(strip=True) if article else ""

    # STOCK
    text = soup.get_text(" ", strip=True)

    if "Нет в наличии" in text:
        stock = "OutOfStock"
    else:
        stock = "InStock"

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

    for a in soup.select(".product-item a[href]"):
        href = a.get("href")

        if href and "/nstrumenti/" in href:
            links.add(href)

    return list(links)


def get_pages(url):
    soup = get_soup(url)
    pages = soup.select(".pagination a.page-link")

    return [p.get("href") for p in pages if p.get("href")]


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
    create_lock()

    if os.path.exists(FILE_PATH):
        os.remove(FILE_PATH)

    try:
        os.makedirs(BASE_DIR, exist_ok=True)

        if not login():
            return

        excel = ExcelWriter(FILE_PATH)

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

            pages = list(set(pages))

            for page in pages:
                print(f"➡ PAGE: {page}")

                product_links = get_products_from_category(page)

                for link in product_links:
                    try:
                        data = parse_product(link)

                        excel.add(
                            data["name"],
                            data["price"],
                            data["article"],
                            data["stock"],
                            data["url"]
                        )
                        excel.save()
                        print("✔", data["name"])

                    except Exception as e:
                        print("❌ product error:", e)

                excel.save()

            save_status({
                "last_category": cat_name,
                "time": time.time()
            })

        print("\n✅ DONE")

    finally:
        remove_lock()


if __name__ == "__main__":
    main()
