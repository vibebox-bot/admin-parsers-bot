import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# =========================
# OUTPUT
# =========================
OUTPUT_DIR = os.path.abspath("output/208")

FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_208_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)

BASE_URL = "https://hi-tech-odessa.com.ua"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()
session.headers.update(HEADERS)

# =========================
# STATUS
# =========================
def save_status(data):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================
# LOAD PAGE (SAFE)
# =========================
def load_page(url):
    try:
        r = session.get(url, timeout=(5, 20))
        if r.status_code != 200:
            return None
        return BeautifulSoup(r.text, "html.parser")
    except:
        return None


# =========================
# LOCK
# =========================
def create_lock():
    if os.path.exists(LOCK_FILE):
        if time.time() - os.path.getmtime(LOCK_FILE) < 300:
            return False
        os.remove(LOCK_FILE)

    open(LOCK_FILE, "w").close()
    return True


def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# GET ALL CATEGORIES
# =========================
def get_categories():
    url = BASE_URL
    soup = load_page(url)

    if not soup:
        return []

    cats = []

    items = soup.select("ul.products li.product-category a")

    for i in items:
        href = i.get("href")
        if href:
            cats.append(href)

    return cats


# =========================
# GET PRODUCT LINKS (with pagination)
# =========================
def get_products_from_category(cat_url):
    links = []
    page = 1

    while True:
        url = f"{cat_url}&paged={page}"
        soup = load_page(url)

        if not soup:
            break

        items = soup.select("ul.products li.product a")

        if not items:
            break

        for i in items:
            href = i.get("href")
            if href and "?product=" in href:
                links.append(href)

        page += 1

        if page > 50:
            break

    return links


# =========================
# PARSE PRODUCT
# =========================
def parse_product(url):
    soup = load_page(url)

    if not soup:
        return None

    title = soup.select_one(".product_title")
    title = title.text.strip() if title else "-"

    sku = soup.select_one(".sku")
    sku = sku.text.strip() if sku else "-"

    price = soup.select_one(".woocommerce-Price-amount")
    price = price.text.strip() if price else "-"

    stock = "В наличии"
    if soup.select_one(".out-of-stock"):
        stock = "Нет в наличии"

    return {
        "title": title,
        "sku": sku,
        "price": price,
        "stock": stock,
        "url": url
    }


# =========================
# RUN
# =========================
def run():

    print("🚀 START PARSER HI-TECH ALL CATEGORIES")

    if not create_lock():
        print("⛔ ALREADY RUNNING")
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.title = "products"

        ws.append(["Название", "SKU", "Цена", "Наличие", "URL"])

        categories = get_categories()

        print(f"📂 CATEGORIES: {len(categories)}")

        all_products = []

        # =========================
        # COLLECT ALL PRODUCTS
        # =========================
        for cat in categories:
            print("📁 CAT:", cat)
            prods = get_products_from_category(cat)
            all_products.extend(prods)

        all_products = list(set(all_products))

        total = len(all_products)

        print(f"📦 TOTAL PRODUCTS: {total}")

        if total == 0:
            save_status({"running": False, "progress": 100, "done": 0, "total": 0})
            return

        done = 0

        for url in all_products:

            data = parse_product(url)

            if data:
                ws.append([
                    data["title"],
                    data["sku"],
                    data["price"],
                    data["stock"],
                    data["url"]
                ])

            done += 1
            progress = round(done / total * 100, 2)

            print(f"[{progress}%] {url}")

            save_status({
                "running": True,
                "progress": progress,
                "done": done,
                "total": total
            })

            # save periodically
            if done % 10 == 0:
                wb.save(FILE_PATH)

            time.sleep(0.2)

        wb.save(FILE_PATH)

        save_status({
            "running": False,
            "progress": 100,
            "done": done,
            "total": total
        })

        print("✅ DONE")

    finally:
        remove_lock()


# =========================
# ENTRY
# =========================
def run_parser():
    run()
