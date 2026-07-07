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

print("🔥 Харьковская Smart-noni")

BASE = "https://daikens.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

EMAIL = "finik257@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/Smart-noni")
FILE_PATH = os.path.join(OUTPUT_DIR, "Smart-noni_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

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
# LOGIN
# =========================
def login():

    session.get(BASE)

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    r = session.post(
        BASE + "/index.php?route=account/login",
        data=payload,
        headers={
            "Referer": BASE + "/index.php?route=account/login"
        },
        allow_redirects=True
    )

    print("LOGIN:", r.status_code)

    check = session.get(BASE + "/index.php?route=account/account")

    if "Моя інформація" in check.text or "Вихід" in check.text:
        print("✅ LOGIN OK")
    else:
        print("❌ LOGIN FAIL")
        
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

    for a in soup.select("#menu-list > li > a.dropdown-img"):    

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE + "/" + href.lstrip("/")

        cats.append({
            "name": clean(a.get_text()),
            "url": href
        })

    return cats


VISITED_CATEGORIES = set()

# =========================
# CATEGORY
# =========================
def parse_category(cat_url):

    if cat_url in VISITED_CATEGORIES:
        return []

    VISITED_CATEGORIES.add(cat_url)

    print()
    print("📂", cat_url)

    result = []
    seen_products = set()

    # =====================================
    # ОБХОД ТОВАРОВ
    # =====================================

    page = 1

    while True:

        if page == 1:
            url = cat_url
        else:
            sep = "&" if "?" in cat_url else "?"
            url = f"{cat_url}{sep}page={page}"

        print(f"📄 PAGE {page}")

        soup = get_soup(url)

        products = soup.select(".product-name a")

        if not products:
            break

        print("FOUND:", len(products))

        added = 0

        for a in products:

            href = a.get("href", "").strip()

            if not href:
                continue

            if not href.startswith("http"):
                href = BASE + "/" + href.lstrip("/")

            if href in seen_products:
                continue

            seen_products.add(href)

            result.append(parse_product(href))
            added += 1

            time.sleep(0.05)

        if added == 0:
            break

        page += 1

    # =====================================
    # ИЩЕМ ДОЧЕРНИЕ КАТЕГОРИИ
    # =====================================

    soup = get_soup(cat_url)

    subcats = []

    for a in soup.select(".thumbnail.subcategory a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE + "/" + href.lstrip("/")

        if href == cat_url:
            continue

        if href in VISITED_CATEGORIES:
            continue

        subcats.append(href)

    print("SUBCATS:", len(subcats))

    for href in subcats:
        result.extend(parse_category(href))

    return result



def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================
    title = ""

    h1 = soup.select_one("h1[itemprop='name']")

    if h1:
        title = clean(h1.get_text())

    # =========================
    # SKU
    # =========================
    sku = ""

    model = soup.select_one("[itemprop='model']")

    if model:
        sku = clean(model.get_text())

    # =========================
    # PRICE
    # =========================
    price = ""

    p = soup.select_one(".autocalc-product-price")

    if p:
        price = clean(p.get_text())

    # =========================
    # STATUS
    # =========================
    status = ""
    
    btn = soup.select_one("#button-cart")
    
    if btn:
        status = clean(btn.get_text())
    

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

        login()

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

                #key = sku if sku else url
                #key = (title, price)

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

        print("✅ Готово. Харьковская Smart-noni")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()


