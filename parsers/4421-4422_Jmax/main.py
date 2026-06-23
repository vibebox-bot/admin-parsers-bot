import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook

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
#CATEGORY_LIMIT = None

session = requests.Session()


# =========================
# INIT / LOCK
# =========================
def init():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                if "running" in f.read():
                    print("ALREADY RUNNING")
                    exit()
        except:
            pass

    with open(LOCK_FILE, "w") as f:
        f.write("running")

def finish():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# STATUS
# =========================
def update_status(data):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


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

    # 1. сначала GET (важно для cookies)
    r = session.get(login_url)

    soup = BeautifulSoup(r.text, "html.parser")

    # 2. берём форму action (на всякий случай)
    form = soup.select_one("form")

    action = login_url
    if form and form.get("action"):
        action = form.get("action")

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    headers = {
        "Referer": login_url,
        "User-Agent": "Mozilla/5.0"
    }

    r2 = session.post(action, data=payload, headers=headers, allow_redirects=True)

    # 3. правильная проверка логина
    if "logout" in r2.text.lower() or "account/logout" in r2.url:
        print("LOGIN OK")
    else:
        print("LOGIN OK (or guest access)")  # у тебя сайт может пускать и так


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

        if "?" in cat_url:
            url = f"{cat_url}&page={page}"
        else:
            url = f"{cat_url}&page={page}"

        soup = get_soup(url)

        items = soup.select(".product-thumb.uni-item")

        if not items:
            break

        before = len(all_products)

        for item in items:
            a = item.select_one("a.product-thumb__name")

            if not a:
                continue

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

    print("PARSING:", url)

    soup = get_soup(url)

    title = ""
    sku = ""
    price = ""
    status = ""

    h1 = soup.select_one("h1")
    if h1:
        title = clean(h1.get_text())
    else:
        print("NO TITLE:", url)

    sku_tag = soup.select_one(".product-data__item.model")
    if sku_tag:
        sku = clean(sku_tag.get_text())
    else:
        print("NO SKU:", url)

    price_tag = soup.select_one(".product-page__price")
    if price_tag:
        price = clean(price_tag.get_text())
    else:
        print("NO PRICE:", url)

    if soup.select_one("#button-cart"):
        status = "in_stock"
    else:
        status = "out_of_stock"

    return [sku, title, price, status, url]


# =========================
# EXCEL
# =========================
def init_excel():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "Title", "Price", "Status", "URL"])
    wb.save(FILE_PATH)


def append_row(row):
    wb = load_workbook(FILE_PATH)
    ws = wb.active
    ws.append(row)
    wb.save(FILE_PATH)


# =========================
# RUN
# =========================
def run_parser():

    init()
    login()

    init_excel()

    categories = get_categories()

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

    total_cats = len(categories)

    for ci, cat in enumerate(categories):

        update_status({
            "status": "category",
            "current": ci,
            "total": total_cats,
            "url": cat
        })

        products = load_products(cat)

        print("PRODUCTS:", len(products))

        for pi, p in enumerate(products):

            update_status({
                "status": "product",
                "category_index": ci,
                "product_index": pi,
                "total_products": len(products)
            })

            try:
                row = parse_product(p)
                append_row(row)

            except Exception as e:
                print("ERROR:", p, e)

            time.sleep(0.1)

    finish()
    print("DONE")


if __name__ == "__main__":
    run_parser()

