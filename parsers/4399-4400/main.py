import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

print("🔥 JUMPEX CLEAN PARSER (NO PLAYWRIGHT)")

BASE = "https://jumpex.com.ua"

# =========================
# ⚙️ SWITCH (1 category / all)
# =========================
CATEGORY_LIMIT = 1   # 👈 ТЕСТ РЕЖИМ
# CATEGORY_LIMIT = None  # 👈 ВСЕ КАТЕГОРИИ

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


# =========================
# STATUS
# =========================

def set_status():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": True,
            "time": datetime.now().strftime("%d.%m %H:%M")
        }, f, ensure_ascii=False, indent=2)


def update_progress(p):
    try:
        if not os.path.exists(STATUS_PATH):
            return

        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = p

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


def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""


# =========================
# CATEGORIES
# =========================

def get_categories():
    soup = get_soup(BASE)

    cats = set()

    for a in soup.select("a[href*='instrumenty']"):
        href = a.get("href")
        if href:
            if href.startswith("/"):
                href = BASE + href
            cats.add(href)

    return list(cats)


# =========================
# PAGINATION FIXED
# =========================

def load_products(cat_url):

    print("CATEGORY:", cat_url)

    all_products = set()
    page = 0

    while True:

        url = f"{cat_url}?start={page}"
        print("LOAD:", url)

        soup = get_soup(url)

        items = soup.select("div.product")

        print("PRODUCTS:", len(items))

        if not items:
            break

        before = len(all_products)

        for item in items:
            a = item.select_one(".name a")
            if not a:
                continue

            href = a.get("href")

            if href.startswith("/"):
                href = BASE + href

            all_products.add(href)

        # STOP CONDITIONS
        if len(all_products) == before:
            break

        if len(items) < 12:
            break

        page += 12
        time.sleep(0.7)

    print("TOTAL LINKS:", len(all_products))
    return list(all_products)


# =========================
# PRODUCT PARSE
# =========================

def parse_product(url):

    soup = get_soup(url)

    title = clean(soup.select_one("h1").get_text()) if soup.select_one("h1") else ""

    sku = ""
    s = soup.select_one(".prod-ean")
    if s:
        sku = clean(s.get_text()).replace("Артикул:", "").strip()

    price = ""
    p = soup.select_one(".prod_price")
    if p:
        price = clean(p.get_text())

    return [sku, title, price, url]


# =========================
# MAIN
# =========================

def run_parser():

    set_status()

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "TITLE", "PRICE", "URL"])

    seen = set()

    cats = get_categories()

    print("CATEGORIES:", len(cats))

    if CATEGORY_LIMIT is not None:
        cats = cats[:CATEGORY_LIMIT]

    total = len(cats)

    all_count = 0

    for i, cat in enumerate(cats, 1):

        update_progress(int(i / total * 100))

        products = load_products(cat)

        print("TOTAL LINKS:", len(products))

        for url in products:

            try:
                data = parse_product(url)

                # =========================
                # FIXED DEDUP LOGIC
                # =========================
                key = data[0] if data[0] else data[3]  # SKU или URL

                if key in seen:
                    continue

                seen.add(key)

                # пропуск пустых названий
                if not data[1]:
                    continue

                ws.append(data)

                all_count += 1

                print("ADDED:", data[1])

            except Exception as e:
                print("ERROR:", e)

            time.sleep(0.2)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(FILE_PATH)

    print("DONE:", all_count)

if __name__ == "__main__":
    run_parser()
