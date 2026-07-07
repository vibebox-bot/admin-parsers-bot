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

print("🔥 Харьковская 208")

BASE = "https://hi-tech-odessa.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/208")
FILE_PATH = os.path.join(OUTPUT_DIR, "208_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "208.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "208.txt")

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

            r = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            #print("URL:", url)
            #print("STATUS:", r.status_code)
            #print("FINAL :", r.url)

            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

        except Exception as e:
            print(e)

        time.sleep(1)

    return BeautifulSoup("", "html.parser")

def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""

def get_categories():

    soup = get_soup(BASE + "/?post_type=product")

    cats = []
    seen = set()

    for a in soup.select("li.product-category a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE.rstrip("/") + "/" + href.lstrip("/")

        if href in seen:
            continue

        seen.add(href)

        name = clean(a.select_one("h2").get_text()) if a.select_one("h2") else clean(a.get_text())

        cats.append({
            "name": name,
            "url": href
        })

    print(f"📂 Найдено категорий: {len(cats)}")

    return cats
    
VISITED_CATEGORIES = set()

def parse_category(cat_url):

    if cat_url in VISITED_CATEGORIES:
        return []

    VISITED_CATEGORIES.add(cat_url)

    result = []

    soup = get_soup(cat_url)

    # =========================
    # Подкатегории
    # =========================

    subcats = soup.select("li.product-category > a")

    if subcats:

        for a in subcats:

            href = a.get("href", "").strip()

            if not href:
                continue

            if not href.startswith("http"):
                href = BASE.rstrip("/") + "/" + href.lstrip("/")

            result.extend(parse_category(href))

        return result

    # =========================
    # Товары
    # =========================

    page = 1
    url = cat_url

    while url:

        soup = get_soup(url)

        products = []

        for li in soup.select("li.product"):

            a = li.find("a", href=True)

            if not a:
                continue

            href = a["href"].strip()

            if not href.startswith("http"):
                href = BASE.rstrip("/") + "/" + href.lstrip("/")

            products.append(href)

        #print(f"📄 Страница {page}: {len(products)} товаров")

        if not products:
            break

        for href in products:

            result.append(parse_product(href))
            time.sleep(0.05)

        next_btn = soup.select_one("a.next.page-numbers")

        if next_btn:

            next_url = next_btn.get("href", "").strip()

            if next_url:

                if not next_url.startswith("http"):
                    next_url = BASE.rstrip("/") + "/" + next_url.lstrip("/")

                #print("➡ Следующая:", next_url)

                url = next_url
                page += 1
                continue

        print("🏁 Последняя страница")
        break

    print(f"✅ Всего в категории: {len(result)}")

    return result
    
    

def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================

    title = clean(
        soup.select_one("h1.product_title").get_text()
    ) if soup.select_one("h1.product_title") else ""

    # =========================
    # SKU
    # =========================

    sku = clean(
        soup.select_one("span.sku").get_text()
    ) if soup.select_one("span.sku") else ""

    # =========================
    # PRICE
    # =========================
    
    price = ""
    
    # Сначала ищем акционную цену
    p = soup.select_one("p.price ins .woocommerce-Price-amount")
    
    if p:
        price = clean(p.get_text())
    
    else:
        p = soup.select_one("p.price .woocommerce-Price-amount")
    
        if p:
            price = clean(p.get_text())

    # =========================
    # STATUS
    # =========================

    if soup.select_one("button.single_add_to_cart_button"):
        status = "В наличии"

    elif soup.select_one("p.stock.out-of-stock"):
        status = "Нет в наличии"

    else:
        status = ""

    return [
        sku,
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
        ws.append(["SKU", "TITLE", "PRICE", "STATUS", "URL"])

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

            for sku, title, price, status, url in items:
                
                #key = url
                
                #if key in seen:
                    #continue
                
                #seen.add(key)

                if not title:
                    continue

                ws.append([
                    sku,
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

        print("✅ Готово. Харьковская 208")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
