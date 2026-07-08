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

print("🔥 Харьковская 4399-4400")

BASE = "https://jumpex.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/138.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ru,en-US;q=0.9,en;q=0.8",
    "Connection": "keep-alive",
    "Referer": BASE + "/",
    "Upgrade-Insecure-Requests": "1"
}

session = requests.Session()
session.headers.update(HEADERS)
session.get(BASE, timeout=30)

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

    session.cookies.clear()

    for _ in range(3):

        try:

            r = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            #print("URL:", url)
            #print("STATUS:", r.status_code)
            #print("FINAL :", r.url)

            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

        except Exception as e:
            print(e)

        time.sleep(1)

    return BeautifulSoup("", "html.parser")

def login():

    print("🔐 LOGIN...")

    login_url = BASE + "/login"

    soup = get_soup(login_url)

    form = soup.select_one("form")

    if not form:
        print("❌ LOGIN FORM NOT FOUND")
        return False

    payload = {}

    for inp in form.select("input"):

        name = inp.get("name")

        if not name:
            continue

        payload[name] = inp.get("value", "")

    payload["username"] = EMAIL
    payload["passwd"] = PASSWORD
    payload["remember"] = "yes"

    action = form.get("action") or "/user/loginsave"

    if not action.startswith("http"):
        action = BASE + action

    r = session.post(
        action,
        data=payload,
        allow_redirects=True,
        timeout=30
    )

    if "/login" not in r.url:
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

    for a in soup.select(".catalog_treenameClass a[href]"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith("/"):
            href = BASE + href

        # пропускаем служебные ссылки
        if href == BASE:
            continue

        if href in seen:
            continue

        seen.add(href)

        cats.append({
            "name": clean(a.get_text()),
            "url": href
        })

    print(f"📂 Найдено категорий: {len(cats)}")

    return cats    

def parse_category(cat_url):

    result = []

    page = 0

    while True:

        url = f"{cat_url}?start={page}"

        soup = get_soup(url)

        products = []

        for a in soup.select("div.product .name a"):

            href = a.get("href", "").strip()

            if not href:
                continue

            if href.startswith("/"):
                href = BASE + href

            if href not in products:
                products.append(href)

        if not products:
            break

        print(f"📄 Страница {page // 12 + 1}: {len(products)} товаров")

        for href in products:
            result.append(parse_product(href))
            time.sleep(0.05)

        # последняя страница
        if len(products) < 12:
            break

        page += 12

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

    sku_tag = soup.select_one(".prod-ean")

    if sku_tag:
        sku = clean(
            sku_tag.get_text().replace("Артикул:", "")
        )

    # =========================
    # PRICE
    # =========================
    price = ""

    p = soup.select_one(".prod_price")

    if p:
        price = clean(p.get_text())

    # =========================
    # STATUS
    # =========================
    status = ""

    s = soup.select_one(".avail")

    if s:
        status = clean(s.get_text())
    else:
        s = soup.select_one(".prod-not-avail")
        if s:
            status = clean(s.get_text())

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

            for sku, title, price, status, url in items:
                
                key = sku if sku else url

                if key in seen:
                    print("🔁 ДУБЛЬ:", key, "|", title)
                    continue
                
                seen.add(key)
                
                if not title:
                    print("❌ Пустой TITLE:", url)
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

        print("✅ Готово. Харьковская 4399-4400")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
