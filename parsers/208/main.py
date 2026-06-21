import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
import time
import random
import os
import json
from datetime import datetime

BASE = "https://hi-tech-odessa.com.ua"

# =========================
# PATHS
# =========================

OUTPUT_DIR = os.path.abspath("output/208")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Харьковская_208_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "lock.txt"
)



HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/137.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "Connection": "keep-alive",
}

# =========================
# STATUS
# =========================

def set_status(running: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "running": running,
                "progress": 0,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M")
            },
            f,
            ensure_ascii=False,
            indent=2
        )

def set_lock(state: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if state:
        with open(LOCK_FILE, "w") as f:
            f.write("running")
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)

def is_locked():
    return os.path.exists(LOCK_FILE)

def update_progress(percent):
    try:

        if not os.path.exists(STATUS_PATH):
            return

        with open(
            STATUS_PATH,
            "r",
            encoding="utf-8"
        ) as f:
            data = json.load(f)

        data["progress"] = percent

        with open(
            STATUS_PATH,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )

    except:
        pass


session = requests.Session()
session.headers.update(HEADERS)


# -------------------------
# SAFE REQUEST
# -------------------------
def get(url):
    global session

    for i in range(5):

        try:
            r = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            if r.status_code == 200:
                return r

            if r.status_code == 404:
                return r

            print("HTTP", r.status_code, url)

            if r.status_code in [400, 403, 429, 500, 502]:
                session.close()

                session = requests.Session()
                session.headers.update(HEADERS)

                time.sleep(3)
                continue

        except Exception as e:
            print("ERROR:", url)
            print(e)

        time.sleep(2)

    return None


def get_categories():
    url = f"{BASE}/?post_type=product"
    r = get(url)
    if not r:
        return []

    soup = BeautifulSoup(r.text, "html.parser")

    cats = []

    for a in soup.select("li.product-category a"):
        href = a.get("href")

        if not href:
            continue

        cats.append(href)

    return list(set(cats))


# -------------------------
# PRODUCT PAGE
# -------------------------
def parse_product(url):
    r = get(url)
    if not r:
        return None

    soup = BeautifulSoup(r.text, "html.parser")

    # TITLE
    title = soup.select_one("h1")
    title = title.get_text(strip=True) if title else ""

    # SKU
    sku = soup.select_one(".sku")
    sku = sku.get_text(strip=True) if sku else ""

    # ЦЕНА (правильно WooCommerce)
    price = ""

    price_block = soup.select_one("p.price")

    if price_block:
        # берем ВСЕ цены внутри блока
        amounts = price_block.select("span.woocommerce-Price-amount bdi")

    clean_prices = []

    for a in amounts:
        txt = a.get_text(strip=True)

        # оставляем только норм цены вида $7.20
        if "$" in txt:
            clean_prices.append(txt)

    # берем последнюю актуальную (обычно она и есть текущая)
    if clean_prices:
        price = clean_prices[-1]

    # STOCK
    stock = "В наличии" if soup.select_one("button.single_add_to_cart_button") else "Нет в наличии"

    return {
        "title": title,
        "sku": sku,
        "price": price,
        "stock": stock,
        "url": url
    }

# -------------------------
# CATEGORY PAGINATION
# -------------------------
def parse_category(cat_url):
    page = 1
    products = []

    while True:

        if page == 1:
            url = cat_url
        else:
            if "?" in cat_url:
                url = cat_url + f"&paged={page}"
            else:
                url = cat_url + f"?paged={page}"

        print("PAGE:", url)

        r = get(url)

        if not r:
            print("STOP: request failed")
            break

        soup = BeautifulSoup(r.text, "html.parser")

        # ИЩЕМ ПОДКАТЕГОРИИ
        subcats = []

        for a in soup.select("li.product-category a"):
            href = a.get("href")

            if href:
                subcats.append(href)

        subcats = list(set(subcats))

        # ЕСЛИ ЕСТЬ ПОДКАТЕГОРИИ → ИДЕМ ВНУТРЬ НИХ
        if subcats:

            print("SUBCATEGORIES FOUND:", len(subcats))

            for subcat in subcats:

                print("ENTER SUBCATEGORY:", subcat)

                try:
                    products.extend(parse_category(subcat))
                except Exception as e:
                    print("SUBCAT ERROR:", e)

            return products

        # ИЩЕМ ТОВАРЫ
        links = []

        for a in soup.select(
            "a.woocommerce-LoopProduct-link, a.woocommerce-loop-product__link"
        ):
            href = a.get("href")

            if href:
                links.append(href)

        links = list(set(links))

        if not links:
            print("STOP EMPTY PAGE")
            break

        print("FOUND:", len(links))

        for link in links:

            try:
                p = parse_product(link)

                if p:
                    products.append(p)

            except Exception as e:
                print("ERR:", e)

        page += 1
        time.sleep(0.5)

    return products


# -------------------------
# MAIN (TEST 1 CAT)
# -------------------------
def main():

    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True)

    try:

        wb = Workbook()
        ws = wb.active

        ws.append([
            "Название",
            "SKU",
            "Цена",
            "Наличие",
            "URL"
        ])

        cats = get_categories()

        print("CATEGORIES:", len(cats))

        all_products = []

        total = len(cats)

        for i, cat in enumerate(cats, 1):

            percent = int(i / total * 100)
            update_progress(percent)

            print("CAT:", cat)

            items = parse_category(cat)
            all_products.extend(items)

        print("TOTAL PRODUCTS:", len(all_products))

        for p in all_products:

            ws.append([
                p["title"],
                p["sku"],
                p["price"],
                p["stock"],
                p["url"]
            ])

        update_progress(100)

        wb.save(FILE_PATH)

        print("SAVED:", FILE_PATH)

    finally:

        set_status(False)
        set_lock(False)


if __name__ == "__main__":
    main()
