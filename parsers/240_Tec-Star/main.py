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

print("🔥 Харьковская 240 Tec-Star")

BASE = "https://tec-star.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/240_Tec-Star")
FILE_PATH = os.path.join(OUTPUT_DIR, "240_Tec-Star_LIVE.xlsx")
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

    categories = []

    menu = soup.select_one("ul.sc-megamenu-list")

    if not menu:
        return categories


    def walk(li):

        child = li.select_one(":scope > div.sc-megamenu-child")

        # Есть подкатегории
        if child:

            items = child.select(":scope > ul > li")

            for item in items:
                walk(item)

            return

        # Конечная категория
        a = li.select_one(":scope > a")

        if not a:
            return

        href = a.get("href", "").strip()

        if not href:
            return

        if href.startswith("/"):
            href = BASE + href

        categories.append(href)


    for li in menu.select(":scope > li"):
        walk(li)

    categories = list(dict.fromkeys(categories))

    print(f"📂 Categories: {len(categories)}")

    return categories
    
# =========================
# LAST PAGE DETECTION
# =========================
def get_last_page(soup):

    pages = [1]

    for a in soup.select("ul.pagination a"):

        href = a.get("href", "")

        m = re.search(r"[?&]page=(\d+)", href)

        if m:
            pages.append(int(m.group(1)))

    return max(pages)

# =========================
# PARSE PRODUCT
# =========================
def parse_product(url):

    soup = get_soup(url)

    sku = ""
    title = ""
    price = ""
    status = ""

    # TITLE
    h1 = soup.select_one("h1")

    if h1:
        title = clean(h1.get_text())

    # SKU
    for div in soup.select(".sc-product-info-item"):

        text = clean(div.get_text())

        if "Код товару" in text:
            sku = (
                text
                .replace("Код товару:", "")
                .replace("Код товару", "")
                .strip()
            )
            break

    # PRICE
    price_el = soup.select_one(".sc-module-price")

    if price_el:
        price = clean(price_el.get_text())

    # STATUS (берем текст кнопки как есть)
    btn = soup.select_one("#button-cart .sc-btn-text")

    if btn:
        status = clean(btn.get_text())

    if not status:

        btn = soup.select_one(".sc-stock-notifier-btn .sc-btn-text")

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
# PARSE CATEGORY
# =========================
# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    first_page = get_soup(cat_url)

    last_page = get_last_page(first_page)

    for page in range(1, last_page + 1):

        if page == 1:
            url = cat_url
        else:

            if "?" in cat_url:
                url = f"{cat_url}&page={page}"
            else:
                url = f"{cat_url}?page={page}"

        soup = get_soup(url)

        cards = soup.select("div.product-layout")

        print(f"📄 Page {page}: {len(cards)} products")

        for card in cards:

            a = card.select_one("a.sc-module-title")

            if not a:
                continue

            href = a.get("href", "").strip()

            if not href:
                continue

            if href.startswith("/"):
                href = BASE + href

            item = parse_product(href)

            all_items.append(item)

    return all_items
  
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

            items = parse_category(cat)

            for sku, title, price, status, url in items:

                #key = sku if sku else url
                key = (title, price)

                if key in seen:
                    continue

                seen.add(key)

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

        print("✅ Готово. Харьковская 240 Tec-Star")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
