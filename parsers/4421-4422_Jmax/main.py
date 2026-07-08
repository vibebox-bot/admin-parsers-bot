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

print("🔥 Харьковская 4421-4422 Jmax")

BASE = "https://hi-tech-odessa.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": BASE
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

            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

        except Exception as e:
            print(e)

        time.sleep(1)

    return BeautifulSoup("", "html.parser")


def login():

    print("🔐 LOGIN...")

    login_url = BASE + "/index.php?route=account/login"

    soup = get_soup(login_url)

    form = soup.select_one("form")

    if not form:
        print("❌ LOGIN FORM NOT FOUND")
        return False

    action = form.get("action") or login_url

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    r = session.post(
        action,
        data=payload,
        allow_redirects=True,
        timeout=30
    )

    if (
        "logout" in r.text.lower()
        or "account/logout" in r.text.lower()
        or "account/account" in r.url
    ):
        print("✅ LOGIN OK")
        return True

    print("❌ LOGIN FAILED")
    return False



def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""

def get_categories():

    soup = get_soup(BASE)

    cats = []
    seen = set()

    for a in soup.select("#menu a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if "route=product/category" not in href:
            continue

        if not href.startswith("http"):
            href = BASE + "/" + href.lstrip("/")

        if href in seen:
            continue

        seen.add(href)

        cats.append({
            "name": clean(a.get_text()),
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

    page = 1

    while True:

        url = f"{cat_url}&page={page}"

        soup = get_soup(url)

        products = []

        for a in soup.select(".product-thumb.uni-item a"):

            href = a.get("href", "").strip()

            if not href:
                continue

            if not href.startswith("http"):
                href = BASE + "/" + href.lstrip("/")

            if href not in products:
                products.append(href)

        if not products:
            break

        print(f"📄 Страница {page}: {len(products)} товаров")

        for href in products:
            result.append(parse_product(href))
            time.sleep(0.05)

        page += 1

    print(f"✅ Всего в категории: {len(result)}")

    return result
    

def parse_product(url):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================
    title = ""

    h1 = soup.select_one("h1")

    if h1:
        title = clean(h1.get_text())

    # =========================
    # SKU
    # =========================
    sku = ""

    sku_tag = soup.select_one(".product-data__item.model")

    if sku_tag:
        sku = clean(
            sku_tag.get_text().replace("Код товара:", "")
        )

    # =========================
    # PRICE
    # =========================
    price = ""

    p = soup.select_one(".product-page__price")

    if p:
        price = clean(p.get_text())

    # =========================
    # STATUS
    # =========================
    status = ""

    btn = soup.select_one("#button-cart span")

    if btn:
        status = clean(btn.get_text())
    else:
        status = "Нет кнопки"

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

        if not login():
            print("❌ Не удалось авторизоваться")
            save_status(False, 0, USER, FILE_PATH)
            return

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

        print("✅ Готово. Харьковская 4421-4422 Jmax")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
