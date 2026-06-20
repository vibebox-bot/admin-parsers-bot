import os
import json
import requests
from urllib.parse import urljoin
from bs4 import BeautifulSoup
from openpyxl import Workbook

BASE_URL = "https://www.top-kitchen.com.ua"

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen.xlsx")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

print("🔥 TOP-KITCHEN FIXED PARSER STARTED")


# =========================
# REQUEST
# =========================
def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    print("🌐", url)
    print("📡", r.status_code)
    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORY LINKS (with pagination)
# =========================
def get_category_pages(cat_url):
    soup = get_soup(cat_url)

    pages = {cat_url}

    for a in soup.select(".pagination a"):
        href = a.get("href")
        if href:
            pages.add(urljoin(BASE_URL, href))

    return list(pages)


# =========================
# PRODUCT LINKS
# =========================
def get_products(soup):
    links = set()

    for a in soup.select(".product-thumb__name, .product-thumb__image a"):
        href = a.get("href")
        if not href:
            continue

        full = urljoin(BASE_URL, href)
        links.add(full)

    return list(links)


# =========================
# PRODUCT PARSE
# =========================
def parse_product(url):
    soup = get_soup(url)

    name = soup.select_one(".heading-h1")
    name = name.get_text(strip=True) if name else ""

    code = soup.select_one(".product-data")
    code = code.get_text(strip=True) if code else ""

    price = soup.select_one(".product-page__price.price")
    price = price.get_text(strip=True) if price else ""

    qty = soup.select_one(".qty-indicator__bar")
    qty = qty.get("data-original-title", "") if qty else ""

    return name, code, price, qty


# =========================
# CATEGORY PARSER
# =========================
def parse_category(ws, cat_url):
    print("\n====================")
    print("📂 CATEGORY:", cat_url)
    print("====================")

    pages = get_category_pages(cat_url)

    for page in pages:
        soup = get_soup(page)

        products = get_products(soup)

        print("🧩 PRODUCTS:", len(products))

        for p in products:
            print("➡️ PARSING:", p)

            try:
                name, code, price, qty = parse_product(p)

                ws.append([
                    cat_url,
                    name,
                    code,
                    price,
                    qty,
                    p
                ])
            except Exception as e:
                print("❌ ERROR:", e)


# =========================
# MAIN
# =========================
def run_parser():

    os.makedirs(OUTPUT_DIR, exist_ok=True)

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

    soup = get_soup(BASE_URL)

    categories = []

    for a in soup.select(".list-group__a"):
        href = a.get("href")
        if href:
            categories.append(urljoin(BASE_URL, href))

    print("📦 CATEGORIES:", len(categories))

    TEST_MODE = True  # ← сначала тест

    for cat in categories[:1] if TEST_MODE else categories:
        parse_category(ws, cat)

    wb.save(FILE_PATH)
    print("✅ SAVED:", FILE_PATH)


if __name__ == "__main__":
    run_parser()
