import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import datetime

# =========================
# CONFIG
# =========================
BASE = "http://www.jmaxtvshop.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

CATEGORY_LIMIT = 1

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": BASE
})

wb = None
ws = None


# =========================
# INIT
# =========================
def init():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(LOCK_FILE):
        print("ALREADY RUNNING")
        exit()

    with open(LOCK_FILE, "w") as f:
        f.write("running")


def finish():
    global wb

    if wb:
        wb.save(FILE_PATH)

    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# STATUS
# =========================
def update_status(data):
    os.makedirs(os.path.dirname(STATUS_PATH), exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def set_status(**kwargs):
    base = {
        "running": True,
        "canceled": False,
        "progress": 0,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    }
    base.update(kwargs)
    update_status(base)


def check_cancel():
    if os.path.exists(STATUS_PATH):
        try:
            with open(STATUS_PATH, "r", encoding="utf-8") as f:
                st = json.load(f)
            if st.get("canceled"):
                return True
        except:
            pass
    return False


# =========================
# HELPERS
# =========================
def get_soup(url):
    r = session.get(url, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


def clean(text):
    return text.strip().replace("\n", " ").replace("\t", " ")


# =========================
# LOGIN
# =========================
def login():
    print("LOGIN...")

    login_url = BASE + "/index.php?route=account/login"
    session.get(login_url)

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    r = session.post(login_url, data=payload, allow_redirects=True)

    if "logout" in r.text.lower():
        print("LOGIN OK")
        return True
    else:
        print("LOGIN FAILED")
        return False


# =========================
# CATEGORIES
# =========================
def get_categories():
    soup = get_soup(BASE)

    cats = []
    for a in soup.select("#menu a"):
        href = a.get("href")
        if href and "route=product/category" in href:
            cats.append(href)

    cats = list(dict.fromkeys(cats))
    print("FOUND CATEGORIES:", len(cats))
    return cats


# =========================
# PRODUCTS LIST
# =========================
def load_products(cat_url):
    print("CATEGORY:", cat_url)

    all_products = set()
    page = 1

    while True:
        if check_cancel():
            print("CANCELED")
            return []

        url = f"{cat_url}&page={page}"
        soup = get_soup(url)

        items = soup.select(".product-thumb.uni-item")

        print(f"PAGE {page} -> ITEMS {len(items)}")

        if not items:
            break

        before = len(all_products)

        for item in items:
            a = item.select_one("a")
            if a:
                href = a.get("href")
                if href:
                    all_products.add(href)

        if len(all_products) == before:
            break

        page += 1
        time.sleep(0.3)

    return list(all_products)


# =========================
# PRODUCT PARSER
# =========================
def parse_product(url):
    soup = get_soup(url)

    title = ""
    sku = ""
    price = ""
    status = ""

    h1 = soup.select_one("h1")
    if h1:
        title = clean(h1.get_text())

    sku_tag = soup.select_one(".product-data__item.model")
    if sku_tag:
        sku = clean(sku_tag.get_text().replace("Код товара:", ""))

    price_tag = soup.select_one(".product-page__price")
    if price_tag:
        price = clean(price_tag.get_text())

    status = "in_stock" if soup.select_one("#button-cart") else "out_of_stock"

    return [sku, title, price, status, url]


# =========================
# EXCEL
# =========================
def init_excel():
    global wb, ws

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "Title", "Price", "Status", "URL"])

    wb.save(FILE_PATH)


# =========================
# RUN
# =========================
def run_parser():
    init()

    ok = login()
    if not ok:
        set_status(running=False, progress=0, error="LOGIN FAILED")
        return

    init_excel()

    categories = get_categories()

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

    total_categories = len(categories)
    total_products_global = 0

    # считаем заранее (чтобы был норм прогресс)
    all_products_map = {}

    for cat in categories:
        prods = load_products(cat)
        all_products_map[cat] = prods
        total_products_global += len(prods)

    done_products = 0

    for ci, cat in enumerate(categories):

        products = all_products_map[cat]

        for pi, p in enumerate(products):

            if check_cancel():
                set_status(running=False, canceled=True, progress=0)
                finish()
                return

            try:
                row = parse_product(p)
                ws.append(row)

            except Exception as e:
                print("ERROR:", p, e)

            done_products += 1

            progress = int((done_products / max(total_products_global, 1)) * 100)

            set_status(
                running=True,
                canceled=False,
                progress=progress,
                category_index=ci,
                product_index=pi,
                total_products=total_products_global,
                file_path=FILE_PATH
            )

            print(f"[{done_products}/{total_products_global}] {progress}%")

            time.sleep(0.1)

    set_status(running=False, canceled=False, progress=100)
    finish()

    print("DONE")


if __name__ == "__main__":
    run_parser()
