import os
import json
import time
from datetime import datetime

import requests
from bs4 import BeautifulSoup

from openpyxl import Workbook

print("🔥 TOP-KITCHEN PARSER LOADED")

# =========================
# PATHS
# =========================

BASE_URL = "https://www.top-kitchen.com.ua"

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_LIVE.xlsx")
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
            "time": datetime.now().strftime("%Y-%m-%d %H:%M")
        }, f, ensure_ascii=False, indent=2)


def set_lock(state: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if state:
        with open(LOCK_FILE, "w") as f:
            f.write("running")
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


def is_locked():
    return os.path.exists(LOCK_FILE)


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
# HTTP
# =========================

def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORIES
# =========================

def get_categories():
    soup = get_soup(BASE_URL)

    cats = []
    seen = set()

    main = soup.select(".list-group__a")
    sub = soup.select(".list-group__children-a")

    for el in main + sub:
        href = el.get("href")

        if not href:
            continue

        if "top-kitchen.com.ua" not in href:
            continue

        if href in seen:
            continue

        seen.add(href)
        cats.append(href)

    print("📦 CATEGORIES FOUND:", len(cats))
    return cats


def get_product_links(soup):
    cards = soup.select(".product-thumb a[href], .product-layout a[href]")

    links = []
    seen = set()

    for c in cards:
        href = c.get("href")

        if not href:
            continue

        if not href.endswith(".html"):
            continue

        if href in seen:
            continue

        seen.add(href)
        links.append(href)

    return links


# =========================
# PRODUCT
# =========================

def parse_product(url):
    soup = get_soup(url)

    try:
        name = soup.select_one(".heading-h1 h1").get_text(strip=True)
    except:
        name = ""

    try:
        model = soup.select_one(".product-data__item.model").get_text(strip=True)
        model = model.replace("Код Товара:", "").strip()
    except:
        model = ""

    try:
        sku = soup.select_one(".product-data__item.sku").get_text(strip=True)
        sku = sku.replace("Артикул:", "").strip()
    except:
        sku = ""

    try:
        price = soup.select_one(".product-page__price.price").get_text(strip=True)
    except:
        price = ""

    qty = ""
    el = soup.select_one(".qty-indicator__bar")

    if el:
        qty = el.get("data-original-title", "") or ""
    else:
        el2 = soup.select_one(".qty-indicator")
        if el2:
            qty = el2.get_text(strip=True)

    return name, model, sku, price, qty


# =========================
# CATEGORY
# =========================

def parse_category(url):
    seen = set()
    page = 1

    while True:

        current_url = url if page == 1 else f"{url}?page={page}"
        soup = get_soup(current_url)

        links = get_product_links(soup)

        if not links:
            break

        new_links = []

        for link in links:
            if link not in seen:
                seen.add(link)
                new_links.append(link)

        if not new_links:
            break

        for link in new_links:
            try:
                name, model, sku, price, qty = parse_product(link)

                ws.append([
                    url,
                    name,
                    model,
                    sku,
                    price,
                    qty,
                    link
                ])
            except:
                pass

        page += 1

        if page > 50:
            break


# =========================
# MAIN
# =========================

def run_parser():
    print("🚀 RUN_PARSER STARTED")

    global wb, ws

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True)

    try:

        wb = Workbook()
        ws = wb.active
        ws.title = "top-kitchen"

        ws.append([
            "Category",
            "Name",
            "Model",
            "SKU",
            "Price",
            "Availability",
            "URL"
        ])

        categories = get_categories()
        total = len(categories)

        print("CATEGORIES:", categories)

        for i, cat in enumerate(categories, 1):

            percent = int(i / total * 100)
            update_progress(percent)

            parse_category(cat)

        update_progress(100)
        wb.save(FILE_PATH)

    finally:
        set_status(False)
        set_lock(False)


if __name__ == "__main__":
    run_parser()
