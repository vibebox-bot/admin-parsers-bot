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

print("🔥 Харьковская 237")

BASE = "https://elite-top.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None


OUTPUT_DIR = os.path.abspath("output/237")
FILE_PATH = os.path.join(OUTPUT_DIR, "237_LIVE.xlsx")
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

    for a in soup.select("ul.nav.navbar-nav > li > a.dropdown-img"):

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
            if "?" in cat_url:
                url = f"{cat_url}&page={page}"
            else:
                url = f"{cat_url}/?page={page}"

        soup = get_soup(url)

        products = soup.select(".product-name a")

        if not products:
            break

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

    return result


def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================
    title = ""

    h1 = soup.select_one("h1.h1-prod-name")

    if h1:
        title = clean(h1.get_text())

    # =========================
    # SKU
    # =========================
    sku = ""

    model = soup.select_one("span[itemprop='model']")

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

    # сначала обычная кнопка "В корзину"
    btn = soup.select_one("#button-cart")

    if btn:
        status = clean(btn.get_text())

    # если её нет — ищем кнопку в блоке cart
    if not status:
        btn = soup.select_one(".cart button")

        if btn:
            span = btn.select_one("span")
            if span:
                status = clean(span.get_text())
            else:
                status = clean(btn.get_text())

    # если всё равно пусто — берём любой текст кнопки покупки
    if not status:
        for btn in soup.select("button"):
            txt = clean(btn.get_text())
            if txt:
                status = txt
                break

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

        print("✅ Готово. Харьковская 237")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()


