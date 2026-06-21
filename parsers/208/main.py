import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

OUTPUT_DIR = os.path.abspath("output/208")

FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_208_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)

BASE_URL = "https://hi-tech-odessa.com.ua"

HEADERS = {"User-Agent": "Mozilla/5.0"}

CATEGORY_URL = None  # будет меняться динамически


# =========================
# STATUS
# =========================
def save_status(data):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================
# LOAD
# =========================
def load(url):
    r = requests.get(url, headers=HEADERS, timeout=20)
    return BeautifulSoup(r.text, "html.parser")


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
# GET CATEGORIES
# =========================
def get_categories():
    url = BASE_URL + "/?post_type=product"
    soup = load(url)

    cats = []

    items = soup.select("li.product-category a")

    for i in items:
        href = i.get("href")
        if href:
            cats.append(href)

    return list(set(cats))


# =========================
# YOUR WORKING FUNCTION
# =========================
def get_products():
    links = []
    page = 1

    while True:
        url = CATEGORY_URL + f"&paged={page}"

        try:
            r = requests.get(url, headers=HEADERS, timeout=20)

            if r.status_code == 404:
                break

            soup = BeautifulSoup(r.text, "html.parser")

            items = soup.select("ul.products li.product a")

            if not items:
                break

            for i in items:
                href = i.get("href")

                if href and "?product=" in href:
                    links.append(href)

            print(f"📄 PAGE {page} OK - {len(items)} items")

        except Exception as e:
            print("PAGE ERROR:", e)
            break

        page += 1

        if page > 100:
            break

    return list(set(links))


# =========================
# PRODUCT PARSE
# =========================
def parse_product(url):
    soup = load(url)

    title = soup.select_one(".product_title")
    sku = soup.select_one(".sku")
    price = soup.select_one(".woocommerce-Price-amount")

    stock = "В наличии"
    if soup.select_one(".out-of-stock"):
        stock = "Нет в наличии"

    return {
        "title": title.text.strip() if title else "-",
        "sku": sku.text.strip() if sku else "-",
        "price": price.text.strip() if price else "-",
        "stock": stock,
        "url": url
    }


# =========================
# RUN ALL
# =========================
def run():

    print("🚀 START ALL CATEGORIES PARSER")

    if not create_lock():
        print("⛔ ALREADY RUNNING")
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.append(["Категория", "Название", "SKU", "Цена", "Наличие", "URL"])

        categories = get_categories()

        print(f"📂 CATEGORIES FOUND: {len(categories)}")

        all_products = []

        # =========================
        # LOOP CATEGORIES
        # =========================
        for cat in categories:

            global CATEGORY_URL
            CATEGORY_URL = cat

            print("\n====================")
            print("📁 CATEGORY:", cat)

            links = get_products()

            for l in links:
                all_products.append((cat, l))

        total = len(all_products)

        print(f"\n📦 TOTAL PRODUCTS: {total}")

        done = 0

        for cat, url in all_products:

            data = parse_product(url)

            ws.append([
                cat,
                data["title"],
                data["sku"],
                data["price"],
                data["stock"],
                data["url"]
            ])

            done += 1
            progress = int(done / total * 100)

            print(f"[{progress}%] {data['title']}")

            save_status({
                "running": True,
                "progress": progress,
                "done": done,
                "total": total
            })

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


def run_parser():
    run()
