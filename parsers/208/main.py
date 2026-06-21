import os
import json
import time
import requests
from bs4 import BeautifulSoup
from datetime import datetime

from openpyxl import Workbook


BASE = "https://hi-tech-odessa.com.ua"

OUTPUT_DIR = os.path.abspath("output/hi_tech_test")

FILE_PATH = os.path.join(OUTPUT_DIR, "hi_tech_test.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


# =========================
# STATUS
# =========================
def save_status(data):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================
# REQUEST
# =========================
def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORY (TEST ONLY 1)
# =========================
def get_test_category():
    soup = get_soup(BASE)

    first = soup.select_one("li.product-category a")

    return {
        "name": first.get_text(strip=True),
        "url": first["href"]
    }


# =========================
# PRODUCTS LIST (PAGINATION)
# =========================
def get_products(category_url):
    products = []
    page = 1

    while True:
        url = f"{category_url}&paged={page}"
        soup = get_soup(url)

        items = soup.select("ul.products li a")

        if not items:
            break

        found = 0

        for a in items:
            href = a.get("href")
            if href and "product" in href:
                products.append(href)
                found += 1

        if found == 0:
            break

        page += 1
        time.sleep(0.3)

    return products


# =========================
# PRODUCT PARSER
# =========================
def parse_product(url):
    soup = get_soup(url)

    name = soup.select_one("h1.product_title")
    name = name.get_text(strip=True) if name else ""

    sku = soup.select_one("span.sku")
    sku = sku.get_text(strip=True) if sku else ""

    price = soup.select_one("span.woocommerce-Price-amount")
    price = price.get_text(strip=True) if price else ""

    in_stock = True
    if soup.select_one("p.stock.out-of-stock"):
        in_stock = False

    return {
        "name": name,
        "sku": sku,
        "price": price,
        "in_stock": in_stock,
        "url": url
    }


# =========================
# EXCEL SAVE
# =========================
def save_excel(data):
    wb = Workbook()
    ws = wb.active

    ws.append(["Название", "Артикул", "Цена", "Наличие", "Ссылка"])

    for item in data:
        ws.append([
            item["name"],
            item["sku"],
            item["price"],
            "YES" if item["in_stock"] else "NO",
            item["url"]
        ])

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(FILE_PATH)


# =========================
# MAIN
# =========================
def run():

    save_status({
        "running": True,
        "progress": 0,
        "time": str(datetime.now())
    })

    print("🚀 START HI-TECH TEST PARSER")

    cat = get_test_category()
    print("📂 CATEGORY:", cat["name"])

    products = get_products(cat["url"])
    total = len(products)

    print("📦 PRODUCTS:", total)

    data = []

    for i, url in enumerate(products, start=1):

        try:
            item = parse_product(url)
            item["category"] = cat["name"]
            data.append(item)

            progress = int((i / total) * 100)

            save_status({
                "running": True,
                "progress": progress,
                "time": str(datetime.now())
            })

            print(f"{progress}% {url}")

        except Exception as e:
            print("ERROR:", e)

        time.sleep(0.2)

    save_excel(data)

    save_status({
        "running": False,
        "progress": 100,
        "time": str(datetime.now()),
        "file_path": FILE_PATH
    })

    print("✅ DONE")


if __name__ == "__main__":
    run()
