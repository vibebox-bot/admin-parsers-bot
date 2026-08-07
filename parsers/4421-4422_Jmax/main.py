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

BASE = "http://www.jmaxtvshop.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 1
CATEGORY_LIMIT = None

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

def get_product_fast(product_id):

    url = BASE + f"/index.php?route=product/product&product_id={product_id}"

    try:

        r = session.get(
            url,
            timeout=10,
            allow_redirects=True
        )

        if r.status_code != 200:
            return None

        soup = BeautifulSoup(r.text, "html.parser")

        h1 = soup.select_one("h1")

        if not h1:
            return None

        title = clean(h1.get_text())

        if not title:
            return None

        return parse_product_soup(soup, url)

    except Exception as e:

        print(f"❌ ID {product_id}: {e}")
        return None

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

    collect_product_links(soup)

    categories = []
    seen = set()

    def add_url(url):

        if not url:
            return

        url = url.strip()

        if not url:
            return

        if url.startswith("#"):
            return

        if url.startswith("javascript"):
            return

        if not url.startswith("http"):
            url = BASE + "/" + url.lstrip("/")

        if "route=product/category" not in url:
            return

        if url in seen:
            return

        seen.add(url)

        categories.append({
            "url": url
        })

    # Ищем абсолютно во всех тегах страницы
    for tag in soup.find_all(True):

        add_url(tag.get("href"))
        add_url(tag.get("data-href"))

    print(f"📂 Categories: {len(categories)}")

    return categories
   
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

        collect_product_links(soup)

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

        #print(f"📄 Страница {page}: {len(products)} товаров")

        for href in products:
            result.append(parse_product(href))
            time.sleep(0.05)

        page += 1

    print(f"✅ Всего в категории: {len(result)}")

    return result

def parse_search(search_text):

    print(f"🔎 Поиск на сайте: {search_text}")

    url = (
        BASE +
        "/index.php?route=product/search"
        f"&search={search_text}"
        "&description=true"
    )

    soup = get_soup(url)

    collect_product_links(soup)

    products = []

    for a in soup.select(".product-thumb.uni-item a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE + "/" + href.lstrip("/")

        if "route=product/product" not in href:
            continue

        if href not in products:
            products.append(href)

    print(f"🔎 Найдено по запросу {search_text}: {len(products)}")

    result = []

    for href in products:

        result.append(parse_product(href))

        time.sleep(0.05)

    return result



def parse_product_soup(soup, url):

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


def parse_product(url):

    soup = get_soup(url)

    if not soup.select_one("h1"):
        return ["", "", "", "", url]

    return parse_product_soup(soup, url)

def parse_all_products():

    print("🔍 Поиск товаров по product_id...")

    result = []

    #max_product_id = 20000
    max_product_id = 14100

    for product_id in range(1, max_product_id + 1):

        item = get_product_fast(product_id)

        if not item:
            continue

        sku, title, price, status, url = item

        result.append(item)

        print(
            f"📦 ID {product_id} | SKU {sku} | {title}"
        )

        if product_id % 100 == 0:

            print(
                f"🔎 Проверено ID: {product_id} | "
                f"Найдено: {len(result)}"
            )

    print(
        f"✅ Всего найдено товаров: {len(result)}"
    )

    return result


PRODUCT_LINKS = set()

def collect_product_links(soup):

    for a in soup.find_all("a", href=True):

        href = a["href"]

        if not href.startswith("http"):
            href = BASE + "/" + href.lstrip("/")

        if "route=product/product" in href:

            PRODUCT_LINKS.add(href)



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

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        print("🧪 ТЕСТ PRODUCT_ID 14060")
        
        test_item = get_product_fast(14060)
        
        if test_item:
            print("✅ PRODUCT_ID 14060 НАЙДЕН:")
            print(test_item)
        else:
            print("❌ PRODUCT_ID 14060 НЕ НАЙДЕН")

        

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        items = parse_all_products()

        print(
            f"📦 Записываем в Excel: {len(items)} товаров"
        )

        total = len(items)

        for i, item in enumerate(items, 1):

            ws.append(item)

            if i % 100 == 0:

                progress = int(
                    i / total * 100
                ) if total else 0

                save_status(
                    True,
                    progress,
                    USER,
                    FILE_PATH
                )

                print(
                    f"💾 Записано: {i}/{total}"
                )

        tmp = FILE_PATH + ".tmp"

        wb.save(tmp)

        os.replace(
            tmp,
            FILE_PATH
        )

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        print(
            f"✅ Готово. "
            f"Харьковская 4421-4422 Jmax"
        )

        print(
            f"📊 Всего товаров: {total}"
        )

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
