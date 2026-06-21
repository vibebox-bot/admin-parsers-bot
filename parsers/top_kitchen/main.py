import os
import json
import time
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook

print("🔥 TOP-KITCHEN PARSER FINAL STABLE")

# =========================
# PATHS (НЕ ТРОГАЕМ)
# =========================
OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

BASE_URL = "http://www.top-kitchen.com.ua"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# =========================
# STATUS SYSTEM (ДЛЯ DASHBOARD)
# =========================
def set_status(running=True, progress=0):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def update_progress(percent):
    set_status(True, percent)


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
# HTTP
# =========================
def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)

    print("🌐", url)
    print("📡", r.status_code)

    return BeautifulSoup(r.text, "html.parser")


# =========================
# PAGINATION (ВАЖНО)
# =========================
def get_all_pages(category_url):
    soup = get_soup(category_url)

    pages = [category_url]

    for a in soup.select(".pagination a"):
        href = a.get("href")
        if href:
            full = urljoin(BASE_URL, href)
            if full not in pages:
                pages.append(full)

    print("📄 PAGES FOUND:", len(pages))
    return pages


# =========================
# PRODUCTS LIST
# =========================
def get_product_links(soup):
    links = set()

    for a in soup.select(".product-thumb__name, .product-thumb__image a"):
        href = a.get("href")

        if not href:
            continue

        full = urljoin(BASE_URL, href)
        links.add(full)

    print("🧩 PRODUCTS FOUND:", len(links))
    return list(links)


# =========================
# PRODUCT PARSER (ТВОИ СЕЛЕКТОРЫ)
# =========================
def parse_product(url):
    soup = get_soup(url)

    # =========================
    # НАЗВАНИЕ
    # =========================
    name_el = soup.select_one(".heading-h1")
    name = name_el.get_text(strip=True) if name_el else ""

    # =========================
    # КОД ТОВАРА (MODEL)
    # =========================
    model = ""

    model_el = soup.select_one(".product-data__item.model")
    if model_el:
        model = model_el.get_text(" ", strip=True)
        model = model.replace("Код Товара:", "").strip()

    # =========================
    # АРТИКУЛ (SKU)
    # =========================
    sku = ""

    sku_el = soup.select_one(".product-data__item.sku")
    if sku_el:
        sku = sku_el.get_text(" ", strip=True)
        sku = sku.replace("Артикул:", "").strip()

    # =========================
    # ЦЕНА
    # =========================
    price_el = soup.select_one(".product-page__price")
    price = price_el.get_text(strip=True) if price_el else ""

    # =========================
    # НАЛИЧИЕ
    # =========================
    qty = ""

    qty_el = soup.select_one(".qty-indicator__bar")
    if qty_el:
        qty = qty_el.get_text(strip=True)

        # иногда текст пустой → пробуем tooltip
        if not qty:
            qty = qty_el.get("data-original-title", "").strip()

    return name, model, sku, price, qty, url


# =========================
# CATEGORY PARSER (1 КАТЕГОРИЯ)
# =========================
def parse_category(category_url, ws):

    pages = get_all_pages(category_url)

    seen = set()
    total_pages = len(pages)

    for i, page in enumerate(pages, 1):

        percent = int((i / total_pages) * 100)
        update_progress(percent)

        soup = get_soup(page)
        links = get_product_links(soup)

        for link in links:

            if link in seen:
                continue
        
            seen.add(link)
        
            print("➡️ PARSING:", link)
        
            try:
                name, model, sku, price, qty, url = parse_product(link)


                ws.append([
                    name,
                    model,
                    sku,
                    qty,
                    price,
                    url
                ])
        
            except Exception as e:
                print("❌ ERROR:", e)


# =========================
# MAIN
# =========================
def run_parser():

    if is_locked():
        print("⛔ ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True, 0)

    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        wb = Workbook()
        ws = wb.active
        ws.title = "tv-shop"

        # HEADER (ТО ЧТО ТЕБЕ НАДО)
        ws.append([
            "Название",
            "Код товара",
            "Артикул",
            "Наличие",
            "Цена",
            "Ссылка"
        ])

        category_url = BASE_URL + "/tv-shop"

        print("🚀 START CATEGORY:", category_url)

        parse_category(category_url, ws)

        set_status(True, 100)

        wb.save(FILE_PATH)

        print("✅ DONE:", FILE_PATH)

    finally:
        set_lock(False)
        set_status(False, 100)


if __name__ == "__main__":
    run_parser()
