import os
import json
import time
from datetime import datetime
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

print("🔥 TOP-KITCHEN PARSER LOADED")

# =========================
BASE_URL = "https://www.top-kitchen.com.ua"

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_FINAL.xlsx")

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

        if href:
            cats.append(urljoin(BASE_URL, href))

    print("📦 CATEGORIES:", len(cats))
    return cats


# =========================
# ПАГИНАЦИЯ КАТЕГОРИЙ
# =========================

def get_all_pages(category_url):
    pages = set()

    soup = get_soup(category_url)

    pages.add(category_url)

    # 🔥 берем все ссылки из пагинации
    for a in soup.select(".pagination a"):
        href = a.get("href")

        if href:
            pages.add(urljoin(BASE_URL, href))

    return sorted(list(pages))


# =========================
# ТОВАРЫ НА СТРАНИЦЕ
# =========================

def get_product_links(soup):
    links = set()

    for a in soup.select(".product-thumb__name, .product-thumb__image a"):
        href = a.get("href")

        if href:
            links.add(urljoin(BASE_URL, href))

    return list(links)


# =========================
# ПРОДУКТ
# =========================

def parse_product(url):
    soup = get_soup(url)

    name = soup.select_one(".heading-h1 h1")
    name = name.get_text(strip=True) if name else ""

    model = soup.select_one(".product-data")
    model = model.get_text(" ", strip=True) if model else ""

    price = soup.select_one(".product-page__price.price")
    price = price.get_text(strip=True) if price else ""

    qty = soup.select_one(".qty-indicator__bar")
    qty = qty.get("data-original-title", "") if qty else ""

    return name, model, price, qty


# =========================
# MAIN
# =========================

def run_parser():

    print("🚀 START PARSER")

    wb = Workbook()
    ws = wb.active
    ws.title = "top-kitchen"

    ws.append([
        "Category",
        "Name",
        "Model",
        "Price",
        "Availability",
        "URL"
    ])

    categories = get_categories()

    # 🔥 TEST MODE (можешь убрать [:1] потом)
    categories = categories[:1]

    for cat in categories:

        print("\n====================")
        print("📂 CATEGORY:", cat)
        print("====================\n")

        pages = get_all_pages(cat)

        print("📄 PAGES FOUND:", len(pages))

        for page in pages:

            soup = get_soup(page)

            links = get_product_links(soup)

            print("🧩 PRODUCTS:", len(links))

            for link in links:

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
    print("\n✅ DONE:", FILE_PATH)


if __name__ == "__main__":
    run_parser()
