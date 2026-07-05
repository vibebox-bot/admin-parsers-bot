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

print("🔥 Харьковская 219 Магнит")

BASE = "https://magnitopt.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "785931"

OUTPUT_DIR = os.path.abspath("output/219")
FILE_PATH = os.path.join(OUTPUT_DIR, "219_LIVE.xlsx")
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

    login_url = "https://magnitopt.com.ua/themes/default/ajax/login.php"

    payload = {
        "email_auth": EMAIL,
        "pass_auth": PASSWORD
    }

    r = session.post(login_url, data=payload, headers={
        "X-Requested-With": "XMLHttpRequest",
        "Referer": "https://magnitopt.com.ua/"
    })

    # проверка успеха
    try:
        data = r.json()
        print("LOGIN RESPONSE:", data)
    except:
        print("LOGIN RAW:", r.text)

    check = session.get("https://magnitopt.com.ua/")

    if "logout" in check.text.lower() or "вихід" in check.text.lower():
        print("✅ LOGIN OK")
    else:
        print("⚠ LOGIN CHECK")
        
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

    # MAGNIT FIX: categories may be in different container
    selectors = [
        "ul.firstUl li a",
        "ul.catalog-menu li a",
        "nav ul li a",
        ".catalog-menu a"
    ]

    for sel in selectors:
        links = soup.select(sel)

        if links:
            for a in links:
                href = a.get("href", "").strip()

                if not href:
                    continue

                if href.startswith("/"):
                    href = BASE + href

                cats.append({
                    "name": clean(a.get_text()),
                    "url": href
                })

            break  # нашли — выходим

    print("CATEGORIES FOUND:", len(cats))

    return cats
    
    
# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    soup = get_soup(cat_url)

    rows = soup.select("tr.itemPosition.simple")

    print("FOUND ROWS:", len(rows))

    items = []

    for row in rows:

        # =========================
        # SKU (td_2)
        # =========================
        sku = ""
        sku_el = row.select_one("td.td_2")
        if sku_el:
            sku = clean(sku_el.get_text())

        # =========================
        # TITLE (td_3)
        # =========================
        title = ""
        url_product = ""
        product_id = ""

        title_el = row.select_one("td.td_3 a")

        if title_el:
            title = clean(title_el.get_text())
            product_id = title_el.get("data-id", "").strip()

        # =========================
        # PRICE (td_5 -> first span.bold)
        # =========================
        price = ""
        price_el = row.select_one("td.td_5 span.bold.block")

        if price_el:
            price = clean(price_el.get_text())

        # =========================
        # STATUS (ВАЖНО: логика кнопок)
        # =========================
        status = ""

        if row.select_one("button.notify"):
            status = "Нет в наличии"

        elif row.select_one("div.not-available"):
            status = "Нет в наличии"

        elif row.select_one("div.are-available"):
            status = "В наличии"

        elif row.select_one("button.radButton.sy"):
            status = "В наличии"

        else:
            status = "Неизвестно"

        # =========================
        # URL fallback
        # =========================
        if product_id:
            url_product = f"https://magnitopt.com.ua/?product_id={product_id}"
        else:
            url_product = ""

        items.append([
            sku,
            title,
            price,
            status,
            url_product
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
        print("DEBUG CATS:", cats)

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

        print("✅ Готово. Харьковская 219 Магнит")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()

