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

print("🔥 Харьковская 201 Matrix")

BASE = "https://matrix7km.com"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 1
CATEGORY_LIMIT = None

EMAIL = "finik257@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/201_Matrix")
FILE_PATH = os.path.join(OUTPUT_DIR, "201_Matrix_LIVE.xlsx")
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

    # открываем страницу логина
    session.get(
        BASE + "/ua/index.php?route=account/login"
    )

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    r = session.post(
        BASE + "/ua/index.php?route=account/login",
        data=payload,
        headers={
            "Referer": BASE + "/ua/index.php?route=account/login"
        },
        allow_redirects=True
    )

    print("LOGIN:", r.status_code)

    check = session.get(
        BASE + "/ua/index.php?route=account/account"
    )

    if "Вихід" in check.text or "Особистий кабінет" in check.text:
        print("✅ LOGIN OK")
    else:
        print("❌ LOGIN FAIL")

def switch_currency():

    payload = {
        "code": "USD",
        "redirect": BASE + "/ua"
    }

    r = session.post(
        BASE + "/ua/index.php?route=common/currency/currency",
        data=payload,
        headers={
            "Referer": BASE + "/ua"
        },
        allow_redirects=True
    )

    print("💲 SWITCH USD:", r.status_code)


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

    soup = get_soup(BASE + "/ua")

    cats = []
    seen = set()

    menu = soup.select_one("nav.ds-menu-catalog-inner")

    if not menu:
        print("❌ Каталог не найден")
        return cats

    top_ul = menu.find("ul")

    if not top_ul:
        print("❌ Список категорий не найден")
        return cats

    # Берем только категории первого уровня
    for li in top_ul.find_all("li", recursive=False):

        a = li.find(
            "a",
            class_="ds-menu-maincategories-item-title"
        )

        if not a:
            continue

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE + href

        if href in seen:
            continue

        seen.add(href)

        name = clean(a.get_text())

        #print(name, "->", href)

        cats.append({
            "name": name,
            "url": href
        })

    #print(f"📂 Найдено категорий: {len(cats)}")

    return cats

# =========================
# CATEGORY
# =========================
def parse_category(cat_url):

    result = []
    seen = set()

    page = 1

    while True:

        if page == 1:
            url = cat_url
        else:
            url = f"{cat_url}?page={page}"

        #print(f"📄 {url}")

        soup = get_soup(url)

        products = soup.select(
            ".content-block .ds-module-title"
        )

        if not products:
            break

        added = 0

        for a in products:

            href = a.get("href", "").strip()

            if not href:
                continue

            if not href.startswith("http"):
                href = BASE + href

            if href in seen:
                continue

            seen.add(href)

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

    h1 = soup.select_one("h1")

    if h1:
        title = clean(h1.get_text())

    # =========================
    # SKU
    # =========================
    sku = ""

    for span in soup.select("span"):

        txt = clean(span.get_text())

        if "Код товару" in txt:

            code = span.select_one(".light-text")

            if code:
                sku = clean(code.get_text())

            break

    # =========================
    # PRICE
    # =========================
    price = ""

    p = soup.select_one(".ds-price-new")

    if p:
        price = clean(p.get_text())


    # =========================
    # STATUS
    # =========================
    status = ""

    # Нет в наличии
    btn = soup.find("button", onclick=re.compile(r"octStockNotifier"))

    if btn:
        status = clean(btn.get_text())

    # В наличии
    if not status:
        btn = soup.find("button", id="button-cart")

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
# MAIN
# =========================
def run_parser():

    if is_locked():
        return

    set_lock(True)

    try:

        save_status(True, 0, USER, FILE_PATH)

        login()
        switch_currency()

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

        print("✅ Готово. Харьковская 201 Matrix")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()


