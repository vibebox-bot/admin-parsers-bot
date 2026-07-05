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

OUTPUT_DIR = os.path.abspath("output/D-Top")
FILE_PATH = os.path.join(OUTPUT_DIR, "D-Top_LIVE.xlsx")
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

    for li in soup.select("ul.firstUl li"):
        a = li.find("a")
        if not a:
            continue

        href = a.get("href", "")
        if href.startswith("/"):
            href = BASE + href

        cats.append({
            "name": a.get_text(strip=True),
            "url": href
        })

    return cats
    
# =========================
# LAST PAGE DETECTION
# =========================
def get_last_page(soup):

    pages = []

    for a in soup.select("ul.pagination a"):
        href = a.get("href", "")

        m = re.search(r"[?&]page=(\d+)", href)
        if m:
            pages.append(int(m.group(1)))

    return max(pages) if pages else 1
    
# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    base_url = cat_url

    first_url = base_url + "&limit=100"

    first_page = get_soup(first_url)
    last_page = get_last_page(first_page)

    for page in range(1, last_page + 1):

        if page == 1:
            url = first_url
        else:
            url = f"{base_url}&limit=100&page={page}"

        soup = get_soup(url)

        cards = soup.select("tr.itemPosition.simple")

        for card in cards:

            title = ""
            sku = ""
            price = ""
            status = ""
            url_product = ""
            product_id = ""

            # ======================
            # TITLE
            # ======================
            title_el = card.select_one("td.td_3 a")

            if title_el:
                title = clean(title_el.get_text())

                product_id = title_el.get("data-id", "").strip()

                href = title_el.get("href", "").strip()
                if href and not href.startswith("#"):
                    if href.startswith("/"):
                        url_product = BASE + href
                    else:
                        url_product = href

            # ======================
            # PRICE (USD)
            # ======================
            price_el = card.select_one("td.td_5 .bold.block")

            if price_el:
                price = clean(price_el.get_text())

            # ======================
            # STATUS / AVAILABILITY
            # ======================
            status_el = card.select_one("td.td_4 .in_box_1")

            if status_el:
                status = clean(status_el.get_text())

            # ======================
            # FALLBACK URL (if no href)
            # ======================
            if not url_product and product_id:
                url_product = f"https://magnitopt.com.ua/product/{product_id}"

            all_items.append([
                sku,
                title,
                price,
                status,
                url_product
            ])

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

        print("✅ Готово. Харьковская 219 Магнит")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()

