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

print("🔥 Харьковская 205 AND")

BASE = "https://andopt2.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/205_AND")
FILE_PATH = os.path.join(OUTPUT_DIR, "205_AND_LIVE.xlsx")
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

    #login_url = BASE + "/login-ru"
    login_url = BASE + "/index.php?route=account/login"

    # Получаем страницу логина (куки + возможные hidden поля)
    r = session.get(login_url)
    soup = BeautifulSoup(r.text, "html.parser")

    payload = {}

    form = soup.select_one("form")

    if form:
        for inp in form.select("input"):
            name = inp.get("name")
            if name:
                payload[name] = inp.get("value", "")

    # Подставляем логин
    payload["email"] = EMAIL
    payload["password"] = PASSWORD

    # Отправляем форму
    session.post(
        login_url,
        data=payload,
        headers={
            "Referer": login_url
        },
        allow_redirects=True
    )

    # Можно проверить успешность авторизации
    account = session.get(BASE)

    if "logout" in account.text.lower() or "выход" in account.text.lower():
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

    categories = []

    menu = soup.select_one("#oct-menu-ul")

    if not menu:
        return categories

    # только первый уровень
    for li in menu.find_all("li", recursive=False):

        a = li.select("> a.oct-menu-a")

        if not a:
            continue

        href = a.get("href", "").strip()

        if not href:
            continue

        if href.startswith("/"):
            href = BASE + href

        categories.append(href)

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
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    first_page = get_soup(cat_url)
    
    last_page = get_last_page(first_page)
    
    #print(f"📄 Pages: {last_page}")
    
    for page in range(1, last_page + 1):
    
        if page == 1:
            soup = first_page
        else:
            soup = get_soup(f"{cat_url}?page={page}")
    
        cards = soup.select("div.product-layout")
    
        #print(f"Page {page}: {len(cards)} products")
    
        for card in cards:
    
            title = ""
            sku = ""
            price = ""
            status = ""
            url = ""


            title_el = card.select_one(".us-module-title a")
            
            if title_el:
            
                title = clean(title_el.get_text())
            
                href = title_el.get("href", "").strip()
            
                if href.startswith("/"):
                    href = BASE + href
            
                url = href
 
    
            
            sku_el = card.select_one(".us-product-list-description")

            if sku_el:
                sku = clean(sku_el.get_text())
                sku = sku.replace("Артикул -", "").strip()
            
            price_el = card.select_one(".us-module-price-actual")
    
            if price_el:
                price = clean(price_el.get_text())


            stock = card.select_one(".quantity__in-stock")

            status = clean(stock.get_text()) if stock else ""
    
            all_items.append([
                sku,
                title,
                price,
                status,
                url
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

        print("✅ Готово. Фабричная 626")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()

