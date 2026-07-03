import os
import json
import time
from datetime import datetime

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
from openpyxl import Workbook

print("🔥 TOP-KITCHEN")

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

def set_status(running=True, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path
    }

    print("WRITE STATUS:", STATUS_PATH)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
        

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        # lock старше часа = считаем зависшим
        if age > 3600:
            os.remove(LOCK_FILE)
            return False

        return True

    except:
        return False


def set_lock(state):

    if state:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


# =========================
# HTTP
# =========================
def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)

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
        qty = qty_el.get("data-original-title", "").strip()
    
        if not qty:
            qty = qty_el.get("title", "").strip()
        
    
    return name, model, sku, price, qty, url


# =========================
# CATEGORY PARSER (1 КАТЕГОРИЯ)
# =========================
def parse_category(category_url, ws):

    pages = get_all_pages(category_url)

    seen = set()

    for page in pages:

        soup = get_soup(page)

        links = get_product_links(soup)

        # ❗ ДОБАВЬ ЭТО
        links = [l for l in links if "/product" in l or "/tv-shop" in l or "http" in l]

        for link in links:
            if link in seen:
                continue

            seen.add(link)


            try:
                name, model, sku, price, qty, url = parse_product(link)

                ws.append([name, model, sku, price, qty, url])

            except Exception as e:
                continue



def get_categories():
    soup = get_soup(BASE_URL)
 
    categories = []

    for a in soup.select("#category-module a.list-group__a"):
        href = a.get("href")

        if not href:
            continue

        full = urljoin(BASE_URL, href)

        categories.append(full)

    return categories
    
# =========================
# MAIN
# =========================
def run_parser(user=""):

    if is_locked():
        return

    set_lock(True)

    set_status(
        running=True,
        user=user
    )

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

        
        categories = get_categories()

        
        for cat in categories:
            parse_category(cat, ws)
                

        set_status(
            running=False,
            user=user,
            file_path=FILE_PATH
        )

        wb.save(FILE_PATH)

        print("✅ DONE:", FILE_PATH)

    finally:
        set_lock(False)

# =========================
# ENTRY
# =========================
import sys

def main():
    user = ""

    if len(sys.argv) > 1:
        user = sys.argv[1]

    run_parser(user)
