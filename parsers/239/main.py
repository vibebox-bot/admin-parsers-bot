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

print("🔥 STARTED Masterberg")

BASE = "https://masterberg.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/239")
FILE_PATH = os.path.join(OUTPUT_DIR, "239_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "239.lock")

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

    login_url = BASE + "/login/"

    r = session.get(login_url)

    soup = BeautifulSoup(r.text, "html.parser")

    payload = {}

    for inp in soup.select("form input"):

        name = inp.get("name")

        if name:
            payload[name] = inp.get("value", "")

    payload["email"] = EMAIL
    payload["password"] = PASSWORD

    session.post(
        login_url,
        data=payload,
        headers={
            "Referer": login_url
        }
    )

    #print("STATUS:", r2.status_code)
    #print("RESPONSE:", r2.text[:300])

    # 5. проверка успеха (очень важно)
    #if "error" not in r2.text.lower():
        #print("LOGIN OK")
    #else:
        #print("LOGIN FAILED")

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
# =========================
# CATEGORIES
# =========================
def get_categories():

    soup = get_soup(BASE)

    categories = []

    for a in soup.select("#menu-list a.dropdown-img"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith("javascript"):
            continue

        if href.startswith("http"):
            url = href
        else:
            url = BASE.rstrip("/") + "/" + href.lstrip("/")

        if url not in categories:
            categories.append(url)

    return categories

# =========================
# LAST PAGE
# =========================
def get_last_page(soup):

    pages = [1]

    for a in soup.select("ul.pagination a[href]"):

        href = a.get("href", "")

        m = re.search(r"[?&]page=(\d+)", href)

        if m:
            pages.append(int(m.group(1)))

    return max(pages)

# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    items = []

    first = get_soup(cat_url)

    last_page = get_last_page(first)

    for page in range(1, last_page + 1):

        if page == 1:
            soup = first
        else:

            if "?" in cat_url:
                url = cat_url + f"&page={page}"
            else:
                url = cat_url + f"/?page={page}"

            soup = get_soup(url)

        cards = soup.select("div.product-layout")

        for card in cards:

            title = ""
            sku = ""
            price = ""
            status = ""
            url = ""

            # -----------------
            # TITLE + URL
            # -----------------

            a = card.select_one(".product-name a")

            if a:

                title = clean(a.get_text())

                url = a.get("href", "")

            # -----------------
            # SKU
            # -----------------

            sku_el = card.select_one(".product-model")

            if sku_el:
                sku = clean(sku_el.get_text())

            # -----------------
            # PRICE
            # -----------------

            price_el = card.select_one("p.price span")

            if price_el:
                price = clean(price_el.get_text())

            # -----------------
            # STATUS
            # -----------------

            btn = card.select_one(".actions .cart button span")

            if btn:
                status = clean(btn.get_text())

            items.append([
                sku,
                title,
                price,
                status,
                url
            ])

    return items
    
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

                key = sku if sku else url

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

        print("DONE")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
