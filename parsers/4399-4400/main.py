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
CATEGORY_ONLY = 1  # 👈 ТЕСТ: 1 категория

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")

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
# CATEGORIES (ВАЖНО ИСПРАВЛЕНО)
# =========================

def get_categories():
    soup = get_soup(BASE)

    cats = set()

    # 👉 ОСНОВНОЙ МЕНЮ САЙТА (ВАШ СЛУЧАЙ)
    for a in soup.select(".menu-wrapper a[href]"):
        href = a.get("href")
        if not href:
            continue
        if "category" in href or "instrumenty" in href:
            if href.startswith("http"):
                cats.add(href)
            else:
                cats.add(BASE + href)

    # fallback
    for a in soup.select("a[href*='route=product/category']"):
        href = a.get("href")
        if href:
            cats.add(BASE + "/" + href.lstrip("/"))

    return list(cats)


# =========================
# LOAD ALL PRODUCTS
# =========================

def load_all_products(category_url):
    print("CATEGORY:", category_url)

    all_links = set()
    page = 1

    while True:

        url = f"{category_url}&page={page}" if "?" in category_url else f"{category_url}?page={page}"

        r = requests.get(url, headers=HEADERS, timeout=30)
        soup = BeautifulSoup(r.text, "html.parser")

        links = set()

        # 👉 ПРАВИЛЬНЫЙ СЕЛЕКТОР ТОВАРОВ
        for a in soup.select("a[href*='product_id']"):
            href = a.get("href")
            if not href:
                continue

            if href.startswith("http"):
                links.add(href)
            else:
                links.add(BASE + href)

        print(f"PAGE {page} OK -> {len(links)}")

        if not links:
            break

        all_links.update(links)

        # конец пагинации
        if len(links) < 48:
            break

        page += 1
        time.sleep(0.3)

    return list(all_links)


# =========================
# PARSE PRODUCT (ВАШ HTML)
# =========================

def parse_product(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    soup = BeautifulSoup(r.text, "html.parser")

    title = clean(soup.select_one("h1.ttl.md.mb25").get_text() if soup.select_one("h1.ttl.md.mb25") else "")

    sku = ""
    sku_el = soup.select_one(".prod-ean")
    if sku_el:
        sku = clean(sku_el.get_text()).replace("Артикул:", "").strip()

    price = clean(soup.select_one(".prod_price").get_text() if soup.select_one(".prod_price") else "")

    status = ""
    if soup.select_one(".avail"):
        status = clean(soup.select_one(".avail").get_text())
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

    print("CATEGORIES FOUND:", len(categories))

    # 👇 ТЕСТ 1 КАТЕГОРИЯ
    categories = categories[:CATEGORY_ONLY]

    seen = set()
    total_items = 0

    for i, cat in enumerate(categories, 1):

        update_progress(int(i / len(categories) * 100))

        products = load_all_products(cat)

        for url in products:

            try:
                data = parse_product(url)

                if not data[0]:
                    continue

                if data[0] in seen:
                    continue

                seen.add(data[0])

                ws.append(data)
                total_items += 1

                print(f"[{total_items}] {data[1]} | {data[2]}")

            except Exception as e:
                print("ERROR:", e)

            time.sleep(0.2)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(FILE_PATH)

    print("DONE:", total_items)

    set_status(False)


if __name__ == "__main__":
    run_parser()
