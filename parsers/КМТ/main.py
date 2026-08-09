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

print("🔥 Харьковская КМТ")

BASE = "https://kmt5.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 1
CATEGORY_LIMIT = None

EMAIL = "finik257@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/КМТ")
FILE_PATH = os.path.join(OUTPUT_DIR, "КМТ_LIVE.xlsx")
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

    # получаем PHPSESSID
    session.get(BASE)

    login_url = BASE + "/login/?ajax=1"

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    r = session.post(
        login_url,
        data=payload,
        headers={
            "X-Requested-With": "XMLHttpRequest",
            "Referer": BASE + "/access-denied"
        }
    )

    print("LOGIN:", r.status_code)

    try:
        print(r.json())
    except:
        print(r.text)

    check = session.get(BASE)

    if "logout" in check.text.lower() or "выйти" in check.text.lower() or "личный кабинет" in check.text.lower():
        print("✅ LOGIN OK")
    else:
        print("⚠ LOGIN CHECK")

def set_product_group():
    url = BASE + "/index.php?route=product/category/setGroup"

    try:
        r = session.get(
            url,
            params={"product_group": "1"},
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "Referer": BASE + "/"
            },
            timeout=30
        )

        print("GROUP SKU:", r.status_code, r.text.strip())

        if r.status_code == 200 and "success" in r.text.lower():
            print("✅ Включено: Сгруппировать по SKU")
            return True

        print("⚠ Не удалось подтвердить группировку по SKU")
        return False

    except Exception as e:
        print("❌ Ошибка группировки:", e)
        return False

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
    seen = set()

    menu = soup.select_one("nav.menu-left > ul")

    if not menu:
        print("❌ Меню категорий не найдено")
        return cats

    def walk_menu(ul):

        for li in ul.find_all("li", recursive=False):

            a = li.find("a", href=True, recursive=False)

            if a:

                href = a["href"].strip()

                if href.startswith("/"):
                    href = BASE + href

                if href.startswith(BASE):

                    if href not in seen:

                        seen.add(href)

                        cats.append({
                            "name": clean(a.get_text()),
                            "url": href
                        })

            child_ul = li.find("ul", recursive=False)

            if child_ul:
                walk_menu(child_ul)

    walk_menu(menu)

    print(f"📂 Найдено категорий со всеми уровнями: {len(cats)}")

    return cats
    
    
# =========================
# PARSE CATEGORY
# =========================
def parse_product(url, status):    

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

    # =========================
    # CODE
    # =========================
    code = ""

    box = soup.select_one(".box-card_code")

    if box:

        text = box.get_text(" ", strip=True)

        m = re.search(r"Код товара:\s*(Ц-\d+)", text)
        if m:
            sku = m.group(1)

        m = re.search(r"Код:\s*(\S+)", text)
        if m:
            code = m.group(1)

    # =========================
    # PRICE
    # =========================
    price = ""

    new_price = soup.select_one(".price__new")

    if new_price:
        price = clean(new_price.get_text())
    else:
        box_price = soup.select_one(".box-card_hryvnia")
        if box_price:
            price = clean(box_price.get_text())

    return [
        sku,
        code,
        title,
        price,
        status,
        url
    ]


# =========================
# PARSE CATEGORY RECURSIVELY
# =========================

def parse_category(cat_url, visited_categories=None):

    if visited_categories is None:
        visited_categories = set()

    if cat_url in visited_categories:
        return []

    visited_categories.add(cat_url)

    result = []
    seen_products = set()
    seen_pages = set()

    print("")
    print("=" * 70)
    print("📂 CATEGORY:", cat_url)
    print("=" * 70)

    # ==========================================================
    # 1. ОБХОДИМ СТРАНИЦЫ КАТЕГОРИИ
    # ==========================================================

    page = 1

    while True:

        if page == 1:
            url = cat_url
        else:
            sep = "&" if "?" in cat_url else "?"
            url = f"{cat_url}{sep}page={page}&ajax=1"

        if url in seen_pages:
            break

        seen_pages.add(url)

        print(f"📄 PAGE {page}: {url}")

        soup = get_soup(url)

        cards = soup.select("div.list-catalog_item")

        print(f"   Найдено карточек: {len(cards)}")

        if not cards:
            break

        added = 0

        for card in cards:

            a = card.select_one(".list-catalog_title a")

            if not a:
                continue

            href = a.get("href")

            if not href:
                continue

            href = href.strip()

            if href.startswith("/"):
                href = BASE + href

            # ==================================================
            # ТОВАР / SKU
            # ==================================================

            if href not in seen_products:

                seen_products.add(href)

                status = ""

                label = card.select_one(".product__label")

                if label:
                    status = clean(label.get_text())

                item = parse_product(href, status)

                if item[2]:
                    result.append(item)
                    added += 1

        print(f"   ➕ Добавлено товаров: {added}")

        # ==================================================
        # NEXT PAGE
        # ==================================================

        next_button = soup.select_one(".btn__more a")

        if not next_button:
            print("   ⛔ Следующей страницы нет")
            break

        next_url = next_button.get("href")

        if not next_url:
            break

        if next_url.startswith("/"):
            next_url = BASE + next_url

        # Если следующая страница уже была
        if next_url in seen_pages:
            break

        page += 1

        time.sleep(0.3)

    # ==========================================================
    # 2. ИЩЕМ ВЛОЖЕННЫЕ КАТЕГОРИИ
    # ==========================================================

    print("🔎 Ищем подкатегории...")

    soup = get_soup(cat_url)

    subcategories = []

    # Возможные блоки категорий
    selectors = [
        ".catalog-group a",
        ".catalog-category a",
        ".category-list a",
        ".sub-category a",
        ".categories a",
        ".box-category a"
    ]

    for selector in selectors:

        for a in soup.select(selector):

            href = a.get("href")

            if not href:
                continue

            href = href.strip()

            if href.startswith("/"):
                href = BASE + href

            if not href.startswith(BASE):
                continue

            # Не считаем товар подкатегорией
            if href in seen_products:
                continue

            if href == cat_url:
                continue

            if href not in subcategories:
                subcategories.append(href)

    print(f"   📂 Подкатегорий найдено: {len(subcategories)}")

    # ==========================================================
    # 3. РЕКУРСИЯ
    # ==========================================================

    for sub_url in subcategories:

        if sub_url in visited_categories:
            continue

        print("")
        print("➡️ Переходим в подкатегорию:")
        print(sub_url)

        sub_items = parse_category(
            sub_url,
            visited_categories
        )

        result.extend(sub_items)

        time.sleep(0.3)

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

        login()

        set_product_group()

        wb = Workbook()
        ws = wb.active
        ws.append(["SKU", "CODE", "TITLE", "PRICE", "STATUS", "URL"])

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

            print(f"📦 ИТОГО ПО КАТЕГОРИИ: {len(items)}")

            #print("TOTAL ITEMS:", len(items))

            for sku, code, title, price, status, url in items:

                #key = sku if sku else url
                #key = (title, price)

                #if key in seen:
                    #continue

                #seen.add(key)

                if not title:
                    continue

                ws.append([
                    sku,
                    code,
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

        print("✅ Готово. Харьковская КМТ")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()


