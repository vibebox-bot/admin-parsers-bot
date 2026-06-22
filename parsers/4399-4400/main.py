import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

print("🔥 JUMPEX CLEAN PARSER (NO PLAYWRIGHT + SESSION FIX)")

BASE = "https://jumpex.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1   # тест
# CATEGORY_LIMIT = None  # все категории

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
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

    login_url = BASE + "/index.php?route=account/login"

    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": login_url
    }

    # 1. GET login page
    r = session.get(login_url, headers=headers)
    soup = BeautifulSoup(r.text, "html.parser")

    # 2. собираем форму (ВАЖНО)
    form = soup.select_one("form")

    payload = {}

    if form:
        for inp in form.select("input"):
            name = inp.get("name")
            value = inp.get("value", "")

            if name:
                payload[name] = value

    # 3. добавляем данные
    payload["email"] = EMAIL
    payload["password"] = PASSWORD

    # 4. POST с referer
    r2 = session.post(
        login_url,
        data=payload,
        headers=headers,
        allow_redirects=True
    )

    print("FINAL URL:", r2.url)

    # 5. ЖЁСТКАЯ проверка
    if "logout" in r2.text.lower() or "/logout" in r2.text.lower():
        print("LOGIN SUCCESS")
    else:
        print("LOGIN FAILED (REAL)")

# =========================
# STATUS
# =========================


def set_status():
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": True,
            "time": datetime.now().strftime("%d.%m %H:%M")
        }, f, ensure_ascii=False, indent=2)


def update_progress(p):
    try:
        if not os.path.exists(STATUS_PATH):
            return

        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = p

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass


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

    print("FOUND MAIN CATEGORIES:", len(cats))

    for c in cats:
        print(c)

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
        print("LOAD:", url)

        soup = get_soup(url)

        items = soup.select("div.product")

        print("PRODUCTS:", len(items))

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

    print("TOTAL LINKS:", len(all_products))
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
    print("PARSED:", sku, "|", title, "|", price, "|", status)

    return [sku, title, price, status, url]

# =========================
# MAIN
# =========================

def run_parser():

    set_status()

    login()  # 🔥 ВАЖНО

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

        update_progress(int(i / total * 100))

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

                print("ADDED:", title, "|", price)

            except Exception as e:
                print("ERROR:", e)

            time.sleep(0.2)

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    wb.save(FILE_PATH)

    print("DONE:", all_count)


if __name__ == "__main__":
    run_parser()
