import os
import json
import requests
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import datetime

print("🔥 TOP-KITCHEN FIXED PARSER STARTED")

BASE_URL = "https://www.top-kitchen.com.ua"

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {"User-Agent": "Mozilla/5.0"}


# =========================
# HTTP
# =========================
def get_soup(url):
    url = urljoin(BASE_URL, url)

    r = requests.get(url, headers=HEADERS, timeout=30)

    print("🌐", url)
    print("📡 STATUS:", r.status_code)

    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORIES
# =========================
def get_categories():
    soup = get_soup(BASE_URL)

    links = []
    seen = set()

    for a in soup.select(".list-group__a, .list-group__children-a"):
        href = a.get("href")
        if not href:
            continue

        full = urljoin(BASE_URL, href)

        if full not in seen:
            seen.add(full)
            links.append(full)

    print("📦 CATEGORIES:", len(links))
    return links


# =========================
# PAGINATION (ВАЖНО)
# =========================
def get_all_pages(category_url):
    pages = set()
    pages.add(category_url)

    soup = get_soup(category_url)

    for a in soup.select(".pagination a"):
        href = a.get("href")
        if href:
            pages.add(urljoin(BASE_URL, href))

    return list(pages)


# =========================
# PRODUCTS LINKS
# =========================
def get_product_links(soup):
    links = set()

    for a in soup.select(".product-thumb__name, .product-thumb__image a"):
        href = a.get("href")
        if href:
            links.add(urljoin(BASE_URL, href))

    return list(links)


# =========================
# PRODUCT PARSER (FIXED)
# =========================
def parse_product(url):
    soup = get_soup(url)

    # NAME
    name = ""
    el = soup.select_one("h1, .heading-h1")
    if el:
        name = el.get_text(strip=True)

    # PRODUCT CODE
    code = ""
    el = soup.select_one(".product-data")
    if el:
        code = el.get_text(" ", strip=True)

    # PRICE
    price = ""
    el = soup.select_one(".product-page__price, .price")
    if el:
        price = el.get_text(strip=True)

    # AVAILABILITY
    qty = ""
    el = soup.select_one(".qty-indicator__bar")
    if el:
        qty = el.get("data-original-title", "").strip()

    return name, code, price, qty


# =========================
# CATEGORY PARSER
# =========================
def parse_category(cat_url):
    pages = get_all_pages(cat_url)

    seen_products = set()

    for page in pages:
        soup = get_soup(page)

        product_links = get_product_links(soup)

        for link in product_links:
            if link in seen_products:
                continue

            seen_products.add(link)

            print("➡️ PARSING:", link)

            try:
                name, code, price, qty = parse_product(link)

                ws.append([
                    cat_url,
                    name,
                    code,
                    price,
                    qty,
                    link
                ])

            except Exception as e:
                print("❌ ERROR:", e)


# =========================
# MAIN
# =========================
wb = Workbook()
ws = wb.active
ws.title = "top_kitchen"

ws.append([
    "Category",
    "Name",
    "Code",
    "Price",
    "Availability",
    "URL"
])


def run():
    cats = get_categories()

    # 🔥 TEST MODE (1 category)
    test_cat = cats[0]
    print("🧪 TEST CATEGORY:", test_cat)

    parse_category(test_cat)

    wb.save(FILE_PATH)
    print("✅ SAVED:", FILE_PATH)


if __name__ == "__main__":
    run()
