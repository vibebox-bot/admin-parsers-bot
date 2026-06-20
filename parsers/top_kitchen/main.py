import os
import json
import time
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

print("🔥 TOP-KITCHEN PARSER LOADED")

# =========================
BASE_URL = "https://www.top-kitchen.com.ua"

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {"User-Agent": "Mozilla/5.0"}
# =========================


def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)

    print("🌐", url)
    print("📡 STATUS:", r.status_code)

    return BeautifulSoup(r.text, "html.parser")


# =========================
# КАТЕГОРИИ
# =========================

def get_categories():
    soup = get_soup(BASE_URL)

    cats = []
    for el in soup.select(".list-group__a"):
        href = el.get("href")
        if href and "top-kitchen.com.ua" in href:
            cats.append(href)

    print("📦 CATEGORIES:", len(cats))
    return cats


# =========================
# ТОВАРЫ В КАТЕГОРИИ
# =========================

def get_product_links(soup):
    links = []

    # 🔥 ВАЖНЫЙ ФИКС САЙТА
    for a in soup.select(".product-thumb__name, .product-thumb__image a"):
        href = a.get("href")

        if not href:
            continue

        if href.startswith("/"):
            href = BASE_URL + href

        if href.endswith(".html"):
            links.append(href)

    return list(set(links))


# =========================
# ПРОДУКТ
# =========================

def parse_product(url):
    soup = get_soup(url)

    # NAME
    name = soup.select_one(".heading-h1 h1")
    name = name.get_text(strip=True) if name else ""

    # MODEL / DATA BLOCK
    model = soup.select_one(".product-data")
    model = model.get_text(" ", strip=True) if model else ""

    # PRICE
    price = soup.select_one(".product-page__price.price")
    price = price.get_text(strip=True) if price else ""

    # AVAILABILITY
    qty = soup.select_one(".qty-indicator__bar")
    qty = qty.get("data-original-title", "") if qty else ""

    return name, model, price, qty


# =========================
# ОСНОВНОЙ ПАРСИНГ
# =========================

def run_parser():

    print("🚀 RUN TEST MODE (1 CATEGORY ONLY)")

    wb = Workbook()
    ws = wb.active
    ws.title = "top-kitchen"

    ws.append(["Category", "Name", "Model", "Price", "Availability", "URL"])

    categories = get_categories()

    # 🔥 ТЕСТ: берем только 1 категорию
    categories = categories[:1]

    print("🧪 TEST CATEGORY:", categories)

    for cat in categories:

        soup = get_soup(cat)

        links = get_product_links(soup)

        print("🧩 PRODUCTS FOUND:", len(links))

        for link in links[:20]:  # 🔥 ограничение для теста

            print("➡️ PARSING:", link)

            try:
                name, model, price, qty = parse_product(link)

                ws.append([
                    cat,
                    name,
                    model,
                    price,
                    qty,
                    link
                ])

            except Exception as e:
                print("❌ ERROR:", e)

    wb.save(FILE_PATH)
    print("✅ DONE:", FILE_PATH)


if __name__ == "__main__":
    run_parser()
