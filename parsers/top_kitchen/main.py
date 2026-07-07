import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Top Kithen")

BASE = "http://www.top-kitchen.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "top_kitchen.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "top_kitchen.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()
session.headers.update(HEADERS)

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

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
# STATUS
# =========================
def save_status(running=False, progress=0, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path
    }

    tmp = STATUS_PATH + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, STATUS_PATH)


# =========================
# HTTP
# =========================
def get_soup(url):

    for _ in range(3):

        try:

            r = session.get(url, timeout=30)
            
            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

        except:
            pass

        time.sleep(1)

    return BeautifulSoup("", "html.parser")

def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""

# =========================
# CATEGORIES
# =========================
def get_categories():

    soup = get_soup(BASE)

    cats = []
    seen = set()

    for a in soup.select("#category-module a.list-group__a, #category-module a.list-group__children-a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE.rstrip("/") + "/" + href.lstrip("/")

        if href in seen:
            continue

        seen.add(href)

        cats.append({
            "name": clean(a.get_text()),
            "url": href
        })

    return cats

VISITED_CATEGORIES = set()

def parse_category(cat_url):

    if cat_url in VISITED_CATEGORIES:
        return []

    VISITED_CATEGORIES.add(cat_url)

    result = []
    seen_products = set()

    page = 1

    while True:

        if page == 1:
            url = cat_url
        else:
            sep = "&" if "?" in cat_url else "?"
            url = f"{cat_url}{sep}page={page}"

        #print(f"📄 {url}")

        soup = get_soup(url)

        products = soup.select("a.product-thumb__name")
        #print(f"📦 Найдено товаров: {len(products)}")

        if not products:
            break

        added = 0

        for a in products:

            href = a.get("href", "").strip()

            if not href:
                continue

            if not href.startswith("http"):
                href = BASE.rstrip("/") + "/" + href.lstrip("/")

            if href in seen_products:
                continue

            seen_products.add(href)

            result.append(parse_product(href))
            added += 1

            time.sleep(0.05)

        if added == 0:
            break

        page += 1

    return result


def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================
    title = ""

    h1 = soup.select_one("h1")

    if h1:
        title = clean(h1.get_text())

    # =========================
    # SKU (Код товара)
    # =========================
    sku = ""

    model = soup.select_one(".product-data__item.model")

    if model:
        txt = clean(model.get_text())
        sku = txt.replace("Код Товара:", "").strip()

    # =========================
    # ARTICLE (Артикул)
    # =========================
    article = ""

    art = soup.select_one(".product-data__item.sku")

    if art:
        txt = clean(art.get_text())
        article = txt.replace("Артикул:", "").strip()

    # =========================
    # PRICE
    # =========================
    price = ""

    p = soup.select_one(".product-page__price")

    if p:
        price = clean(p.get_text())

    # =========================
    # STATUS
    # =========================
    status = ""
    
    qty = soup.select_one(".qty-indicator__bar")
    
    if qty:
        status = clean(qty.get("title", ""))
        

    return [
        sku,
        article,
        title,
        price,
        status,
        url
    ]
    
# =========================
# MAIN
# =========================
def run_parser():

    if is_locked():
        return

    set_lock(True)

    try:

        save_status(True, 0, USER, FILE_PATH)

        wb = Workbook()
        ws = wb.active
        ws.append(["SKU", "ARTICLE", "TITLE", "PRICE", "STATUS", "URL"])

        seen = set()

        cats = get_categories()

        #cats = [cats[8]]
        
        #print("DEBUG CATS:", cats)
        print(f"📂 Категорий: {len(cats)}")

        if CATEGORY_LIMIT:
            cats = cats[:CATEGORY_LIMIT]

        total = len(cats)

        if total == 0:
            save_status(False, 100, USER, FILE_PATH)
            return

        for i, cat in enumerate(cats, 1):

            save_status(
                True,
                int(i / total * 100),
                USER,
                FILE_PATH
            )

            items = parse_category(cat["url"])

            #print("TOTAL ITEMS:", len(items))

            for sku, article, title, price, status, url in items:
                
                key = url
                
                if key in seen:
                    continue
                
                seen.add(key)

                if not title:
                    continue

                ws.append([
                    sku,
                    article,
                    title,
                    price,
                    status,
                    url
                ])
                

            time.sleep(0.2)

        os.makedirs(OUTPUT_DIR, exist_ok=True)

        tmp = FILE_PATH + ".tmp"

        wb.save(tmp)

        os.replace(tmp, FILE_PATH)

        save_status(False, 100, USER, FILE_PATH)

        print("✅ Готово. Top Kithen")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
