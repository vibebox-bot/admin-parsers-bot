import os
import json
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

print("🔥 TOP-KITCHEN PARSER LOADED")

# =========================
# CONFIG
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


def update_progress(percent):
    if not os.path.exists(STATUS_PATH):
        return

    try:
        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = percent

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass


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


# =========================
# REQUEST
# =========================

def get_soup(url):
    if url.startswith("/"):
        url = BASE_URL + url

    if not url.startswith("http"):
        url = "https://" + url

    r = requests.get(url, headers=HEADERS, timeout=30)

    print("🌐 URL:", url)
    print("📡 STATUS:", r.status_code)
    print("📦 LEN:", len(r.text))

    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORIES
# =========================

def get_categories():
    soup = get_soup(BASE_URL)

    cats = []
    seen = set()

    for el in soup.select(".list-group__a, .list-group__children-a"):
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


# =========================
# PRODUCTS LINKS
# =========================

def get_product_links(soup):
    print("🔎 SEARCH PRODUCTS ON PAGE")

    links = []
    seen = set()

    # 🔥 ГЛАВНЫЙ ФИКС (РАБОТАЕТ НА ТВОЁМ HTML)
    for a in soup.select(".product-thumb__name"):
        href = a.get("href")

        if not href:
            continue

        if href in seen:
            continue

        seen.add(href)
        links.append(href)

    print("🧩 FOUND LINKS:", len(links))
    return links


# =========================
# PRODUCT PARSER
# =========================

def parse_product(url):
    soup = get_soup(url)

    # NAME
    name = ""
    el = soup.select_one(".heading-h1 h1")
    if el:
        name = el.get_text(strip=True)

    # MODEL / CODE
    model = ""
    el = soup.select_one(".product-data__item.model")
    if el:
        model = el.get_text(strip=True).replace("Код Товара:", "").strip()

    # SKU
    sku = ""
    el = soup.select_one(".product-data__item.sku")
    if el:
        sku = el.get_text(strip=True).replace("Артикул:", "").strip()

    # PRICE
    price = ""
    el = soup.select_one(".product-page__price.price")
    if el:
        price = el.get_text(strip=True)

    # AVAILABILITY
    qty = ""
    el = soup.select_one(".qty-indicator__bar")
    if el:
        qty = el.get("data-original-title", "") or ""
    else:
        el = soup.select_one(".qty-indicator")
        if el:
            qty = el.get_text(strip=True)

    return name, model, sku, price, qty


# =========================
# CATEGORY PARSER (WITH PAGINATION)
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
            print("➡️ PARSING:", link)

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
            except Exception as e:
                print("❌ ERROR:", e)

        page += 1

        if page > 50:
            break


# =========================
# MAIN
# =========================

def run_parser():
    print("🚀 RUN PARSER STARTED")

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

        # 🔥 TEST MODE (1 категория сначала)
        categories = categories[:1]

        total = len(categories)

        print("🧪 TEST CATEGORIES:", categories)

        for i, cat in enumerate(categories, 1):
            percent = int(i / total * 100)
            update_progress(percent)

            print("====================")
            print("📂 CATEGORY:", cat)
            print("====================")

            parse_category(cat)

        update_progress(100)

        wb.save(FILE_PATH)
        print("✅ DONE:", FILE_PATH)

    finally:
        set_status(False)
        set_lock(False)


# =========================
# START
# =========================

if __name__ == "__main__":
    run_parser()
