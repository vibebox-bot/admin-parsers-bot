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

print("🔥 JUMPEX")

BASE = "https://jumpex.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
# CATEGORY_LIMIT = 1   # тест
CATEGORY_LIMIT = None  # все категории

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# =========================
# SESSION (ВАЖНО)
# =========================
session = requests.Session()
session.headers.update(HEADERS)

def login():

    print("LOGIN...")

    login_url = BASE + "/login"

    r = session.get(login_url)

    soup = BeautifulSoup(r.text, "html.parser")

    payload = {}

    for inp in soup.select("form input"):
        name = inp.get("name")

        if not name:
            continue

        payload[name] = inp.get("value", "")

    payload["username"] = EMAIL
    payload["passwd"] = PASSWORD
    payload["remember"] = "yes"

    r = session.post(
        BASE + "/user/loginsave",
        data=payload,
        allow_redirects=True
    )

    #print("FINAL URL:", r.url)

    if "/login" not in r.url:
        print("LOGIN OK")
    else:
        print("LOGIN FAILED")
        #print(r.text[:5000])

    #print(session.cookies.get_dict())


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
    r = session.get(url, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""


# =========================
# CATEGORIES
# =========================

def get_categories():

    soup = get_soup(BASE)

    cats = []

    for a in soup.select(".catalog_treenameClass li.nav-item.parent > a"):

        href = a.get("href")

        if not href:
            continue

        if href.startswith("/"):
            href = BASE + href

        cats.append(href)

    cats = list(dict.fromkeys(cats))

    #print("FOUND MAIN CATEGORIES:", len(cats))

    #for c in cats:
        #print(c)

    return cats


# =========================
# PRODUCTS
# =========================

def load_products(cat_url):

    print("CATEGORY:", cat_url)

    all_products = set()
    page = 0

    while True:

        url = f"{cat_url}?start={page}"
        #print("LOAD:", url)

        soup = get_soup(url)

        items = soup.select("div.product")

        #print("PRODUCTS:", len(items))

        if not items:
            break

        before = len(all_products)

        for item in items:
            a = item.select_one(".name a")
            if not a:
                continue

            href = a.get("href")

            if href.startswith("/"):
                href = BASE + href

            all_products.add(href)

        if len(all_products) == before:
            break

        if len(items) < 12:
            break

        page += 12
        time.sleep(0.5)

    #print("TOTAL LINKS:", len(all_products))
    return list(all_products)


# =========================
# PRODUCT PARSE (ЦЕНА ТЕПЕРЬ РАБОТАЕТ)
# =========================

def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================
    title_tag = soup.select_one("h1")
    title = clean(title_tag.get_text()) if title_tag else ""

    # =========================
    # SKU
    # =========================
    sku_tag = soup.select_one(".prod-ean")
    sku = clean(sku_tag.get_text().replace("Артикул:", "")) if sku_tag else ""

    # =========================
    # PRICE
    # =========================
    
    price = ""
    
    price_tag = soup.select_one(".prod_price")
    
    if price_tag:
        price = clean(price_tag.get_text())
    
    # =========================
    # STATUS
    # =========================
    status_tag = soup.select_one(".avail, .prod-not-avail")
    status = status_tag.get_text(strip=True) if status_tag else ""

    # =========================
    # DEBUG
    # =========================

    #print("COOKIES:", session.cookies.get_dict())
    #print("PARSED:", sku, "|", title, "|", price, "|", status)

    return [sku, title, price, status, url]

# =========================
# MAIN
# =========================

def run_parser():

    save_status(
        running=True,
        progress=0,
        user=USER,
        file_path=FILE_PATH
    )
    
    login()

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "TITLE", "PRICE", "STATUS", "URL"])


    seen = set()

    cats = get_categories()

    print("CATEGORIES:", len(cats))

    if CATEGORY_LIMIT is not None:
        cats = cats[:CATEGORY_LIMIT]

    total = len(cats)
    all_count = 0

    for i, cat in enumerate(cats, 1):

        save_status(
            running=True,
            progress=int(i / total * 100),
            user=USER,
            file_path=FILE_PATH
        )

        products = load_products(cat)

        print("TOTAL LINKS:", len(products))

        for url in products:

            try:
                data = parse_product(url)

                sku = data[0].strip() if data[0] else ""
                title = data[1].strip() if data[1] else ""
                price = data[2].strip() if data[2] else ""
                clean_url = data[3].split("?")[0]

                key = sku if sku else clean_url

                if key in seen:
                    continue

                seen.add(key)

                if not title:
                    continue

                sku, title, price, status, url = data
                ws.append([sku, title, price, status, url])

                all_count += 1

                #print("ADDED:", title, "|", price)

            except Exception as e:
                print("ERROR:", e)

            time.sleep(0.2)

    
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    tmp = FILE_PATH + ".tmp"
    wb.save(tmp)
    os.replace(tmp, FILE_PATH)
    
    save_status(
        running=False,
        progress=100,
        user=USER,
        file_path=FILE_PATH
    )
    
    print("DONE:", all_count)
    

if __name__ == "__main__":
    run_parser()
