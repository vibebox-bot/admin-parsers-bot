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

print("🔥 LunaBag")

BASE = "https://luna-toys.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/LunaBag")
FILE_PATH = os.path.join(OUTPUT_DIR, "LunaBag_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()
session.headers.update(HEADERS)

# Переключаем валюту
try:
    session.get(
        BASE + "/_widget/currency_selector/change/3",
        timeout=30
    )
except:
    pass

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

    csrf = "1bc3361d698d203571a8ef05cbebe4c069d5dead"

    try:

        r = session.post(
            url,
            data={
                "catalogBuilder": "1"
            },
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "X-CSRF-Token": csrf,
                "Referer": url,
            },
            timeout=30
        )

        if r.status_code != 200:
            return BeautifulSoup("", "html.parser")

        data = r.json()

        html = ""

        if "products" in data["response"]["html"]:
            html += data["response"]["html"]["products"]

        if "pagination" in data["response"]["html"]:
            html += data["response"]["html"]["pagination"]

        return BeautifulSoup(html, "html.parser")

    except Exception as e:
        print(e)
        return BeautifulSoup("", "html.parser")


def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""
    
# =========================
# CATEGORIES
# =========================
def get_categories():

    r = session.get(BASE + "/katalog/", timeout=30)
    soup = BeautifulSoup(r.text, "html.parser")

    for a in soup.select("a.productsMenu-submenu-a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith("/"):
            href = BASE + href

        if href in seen:
            continue

        seen.add(href)
        categories.append(href)

    print(f"📂 Categories: {len(categories)}")

    return categories
    
    
def get_pages(cat_url):

    soup = get_soup(cat_url)

    pages = [cat_url]
    seen = {cat_url}

    for a in soup.select("nav.pager a[href]"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if "page=all" in href:
            continue

        if href.startswith("/"):
            href = BASE + href

        if href not in seen:
            seen.add(href)
            pages.append(href)

    return pages

# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    result = []

    pages = get_pages(cat_url)

    for page in pages:

        soup = get_soup(page)

        cards = soup.select("div.catalogCard-box")

        for card in cards:

            title = ""
            sku = ""
            price = ""
            status = ""
            url = ""

            a = card.select_one(".catalogCard-title a")

            if a:
                title = clean(a.get_text())

                href = a.get("href", "").strip()

                if href.startswith("/"):
                    href = BASE + href

                url = href

            sku_el = card.select_one(".catalogCard-code")

            if sku_el:
                sku = clean(
                    sku_el.get_text()
                ).replace("Артикул:", "").strip()

            price_el = card.select_one(".catalogCard-price")

            if price_el:
                price = clean(price_el.get_text())

            status_el = card.select_one(".catalogCard-availability")

            if status_el:
                status = clean(status_el.get_text())

            result.append([
                sku,
                title,
                price,
                status,
                url
            ])

    return result
  
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

        print("✅ Готово. LunaBag")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
