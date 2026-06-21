import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

# =========================
# CONFIG
# =========================

BASE = "https://jumpex.com.ua"
CATEGORY_ONLY = 1  # 👈 ВАЖНО: пока тест 1 категория

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# =========================
# STATUS
# =========================

def set_status(running: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": 0,
            "time": datetime.now().strftime("%d.%m %H:%M")
        }, f, ensure_ascii=False, indent=2)


def update_progress(percent):
    try:
        if not os.path.exists(STATUS_PATH):
            return
        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = percent

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass


# =========================
# HELPERS
# =========================

def clean(x):
    return re.sub(r"\s+", " ", x).strip() if x else ""


def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORIES
# =========================

def get_categories():
    soup = get_soup(BASE)

    cats = []
    for a in soup.select("a[href*='category']"):
        href = a.get("href")
        if href and href.startswith("http"):
            cats.append(href)

    # уник
    return list(dict.fromkeys(cats))


# =========================
# LOAD ALL PRODUCTS (ВАЖНО)
# =========================

def load_all_products(category_url):
    print("CATEGORY:", category_url)

    all_links = set()

    page = 1
    while True:

        url = f"{category_url}?page={page}"
        soup = get_soup(url)

        links = set()

        for a in soup.select("a[href*='product_id']"):
            links.add(a["href"] if a["href"].startswith("http") else BASE + a["href"])

        print(f"PAGE {page} OK -> {len(links)}")

        if not links:
            break

        all_links.update(links)

        # если меньше 48 — значит конец
        if len(links) < 48:
            break

        page += 1
        time.sleep(0.5)

    return list(all_links)


# =========================
# PARSE PRODUCT
# =========================

def parse_product(url):
    soup = get_soup(url)

    title = clean(soup.select_one("h1.ttl").text if soup.select_one("h1.ttl") else "")

    sku = clean(soup.select_one(".prod-ean").text if soup.select_one(".prod-ean") else "")
    sku = sku.replace("Артикул:", "").strip()

    price = clean(soup.select_one(".prod_price").text if soup.select_one(".prod_price") else "")

    status = ""
    if soup.select_one(".avail"):
        status = clean(soup.select_one(".avail").text)
    elif soup.select_one(".prod-not-avail"):
        status = "Розпродано"

    return [sku, title, price, status, url]


# =========================
# MAIN
# =========================

def run_parser():

    set_status(True)

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "Title", "Price", "Status", "URL"])

    categories = get_categories()

    print("CATEGORIES:", len(categories))

    categories = categories[:CATEGORY_ONLY]  # 👈 TEST MODE

    total = len(categories)
    seen = set()
    count = 0

    for i, cat in enumerate(categories, 1):

        update_progress(int(i / total * 100))

        products = load_all_products(cat)

        for url in products:

            try:
                data = parse_product(url)

                if data[0] and data[0] in seen:
                    continue

                if data[0]:
                    seen.add(data[0])

                if not data[1]:
                    continue

                ws.append(data)
                count += 1

                print(f"[{count}] {data[1]} | {data[2]}")

            except Exception as e:
                print("ERROR:", e)

            time.sleep(0.2)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(FILE_PATH)

    print("DONE:", count)
    set_status(False)


if __name__ == "__main__":
    run_parser()
