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

HEADERS = {"User-Agent": "Mozilla/5.0"}

# =========================
# HELPERS
# =========================

def fix_url(url):
    if not url:
        return None
    if url.startswith("http"):
        return url
    return BASE_URL + "/" + url.lstrip("/")


def get_soup(url):
    url = fix_url(url)

    r = requests.get(url, headers=HEADERS, timeout=30)

    print("🌐", url)
    print("📡 STATUS:", r.status_code)

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
        href = fix_url(href)

        if not href:
            continue

        if href in seen:
            continue

        seen.add(href)
        cats.append(href)

    print("📦 CATEGORIES:", len(cats))
    return cats


# =========================
# PAGINATION
# =========================

def get_pages(category_url):
    soup = get_soup(category_url)

    pages = set()
    pages.add(category_url)

    for a in soup.select("ul.pagination a"):
        href = fix_url(a.get("href"))
        if href:
            pages.add(href)

    return list(pages)


# =========================
# PRODUCTS
# =========================

def get_product_links(soup):
    links = set()

    # 🔥 правильный селектор карточек
    for a in soup.select(".product-thumb__name"):
        href = fix_url(a.get("href"))
        if href:
            links.add(href)

    return list(links)


# =========================
# PRODUCT PARSER (FIXED)
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

    # AVAILABILITY (ВАЖНО — твой фикс)
    qty = ""
    el = soup.select_one(".qty-indicator__bar")
    if el:
        qty = el.get("data-original-title", "").strip()

    return name, model, sku, price, qty


# =========================
# CATEGORY PARSER
# =========================

def parse_category(cat_url, ws):
    print("\n====================")
    print("📂 CATEGORY:", cat_url)
    print("====================")

    pages = get_pages(cat_url)

    for page_url in pages:
        soup = get_soup(page_url)

        links = get_product_links(soup)

        print("🧩 PRODUCTS:", len(links))

        for link in links:
            link = fix_url(link)

            print("➡️ PARSING:", link)

            try:
                name, model, sku, price, qty = parse_product(link)

                ws.append([
                    cat_url,
                    name,
                    model,
                    sku,
                    price,
                    qty,
                    link
                ])

            except Exception as e:
                print("❌ ERROR:", e)


# =========================
# MAIN
# =========================

def run():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

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

    # 🔥 ТЕСТ: только 1 категория
    test_categories = categories[:1]

    print("🧪 TEST CATEGORY:", test_categories)

    for cat in test_categories:
        parse_category(cat, ws)

    wb.save(FILE_PATH)

    print("✅ DONE:", FILE_PATH)


if __name__ == "__main__":
    run()
