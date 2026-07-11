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

print("🔥 Харьковская D-Top")

BASE = "http://www.dtopelectronic.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 2
#CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

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

    categories = set()
    checked = set()

    def scan(url):

        if url in checked:
            return

        checked.add(url)

        soup = get_soup(url)

        # все ссылки меню
        for a in soup.select("a[href]"):

            href = a.get("href", "").strip()

            if not href:
                continue

            if href.startswith("javascript"):
                continue

            if href.startswith("#"):
                continue

            if href.startswith("/"):
                href = BASE + href

            if not href.startswith(BASE):
                continue

            # только ссылки категорий OpenCart
            if "route=product/category" not in href:
                continue

            # убираем limit/page
            href = re.sub(r'([?&])page=\d+', '', href)
            href = re.sub(r'([?&])limit=\d+', '', href)
            href = href.rstrip("&?")

            if href not in categories:
                categories.add(href)
                #print("📂", href)

                # идём глубже
                scan(href)

    scan(BASE)

    categories = sorted(categories)

    print(f"📂 TOTAL CATEGORIES: {len(categories)}")

    return categories
    
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

    if not pages:
        return 1

    return max(pages)
    
# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    base_url = cat_url
    
    first_page = get_soup(base_url)
    
    last_page = get_last_page(first_page)

    print(cat_url)
    print("LAST PAGE =", last_page)
    
    for page in range(1, last_page + 1):
    
        if page == 1:
            url = base_url
        else:
            url = f"{base_url}&page={page}"
      

        soup = get_soup(url)

        cards = soup.select("div.product-thumb")
        print("PAGE", page, "PRODUCTS", len(cards))

        #print(f"Page {page}: {len(cards)} products")

        for card in cards:

            title = ""
            sku = ""
            price = ""
            status = ""
            url_product = ""

            # TITLE + URL
            title_el = card.select_one(".product-thumb__name")

            if title_el:
                title = clean(title_el.get_text())

                href = title_el.get("href", "").strip()

                if href.startswith("/"):
                    href = BASE + href

                url_product = href

            # SKU
            sku_el = card.select_one(".product-thumb__model")

            if sku_el:
                sku = clean(sku_el.get_text())
                sku = sku.replace("Код товара:", "").strip()

            # PRICE
            price_el = card.select_one(".product-thumb__price")

            if price_el:
                price = clean(price_el.get_text())

            # STATUS (ВАЖНО — БЕЗ ГАДАНИЯ)
            status_el = card.select_one(".qty-indicator__text")

            status = clean(status_el.get_text()) if status_el else ""

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

        print("✅ Готово. Харьковская D-Top")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()

