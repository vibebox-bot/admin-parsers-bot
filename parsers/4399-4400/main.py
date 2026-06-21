import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from urllib.parse import urljoin
from datetime import datetime

BASE = "https://jumpex.com.ua"

LOGIN = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)

session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})


# =========================
# STATUS
# =========================
def set_status(running):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": 0,
            "time": str(datetime.now())
        }, f, ensure_ascii=False, indent=2)


def update_progress(p):
    if not os.path.exists(STATUS_PATH):
        return
    with open(STATUS_PATH, "r", encoding="utf-8") as f:
        data = json.load(f)

    data["progress"] = p

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def set_lock(state):
    if state:
        open(LOCK_FILE, "w").write("1")
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


def is_locked():
    return os.path.exists(LOCK_FILE)


# =========================
# LOGIN
# =========================
def login():
    r = session.get(BASE + "/login")
    soup = BeautifulSoup(r.text, "html.parser")

    data = {
        "username": LOGIN,
        "passwd": PASSWORD
    }

    hidden = soup.select("input[type=hidden]")
    for h in hidden:
        if h.get("name"):
            data[h["name"]] = h.get("value", "1")

    r = session.post(BASE + "/user/loginsave", data=data)

    print("LOGIN:", r.status_code)


# =========================
# CATEGORY (1 TEST)
# =========================
def get_category():
    r = session.get(BASE)
    soup = BeautifulSoup(r.text, "html.parser")

    a = soup.select_one("li.nav-item.parent a[href]")

    if not a:
        raise Exception("NO CATEGORY")

    return urljoin(BASE, a["href"])


# =========================
# PRODUCTS FROM CATEGORY
# =========================
def get_products(cat_url):

    r = session.get(cat_url)
    soup = BeautifulSoup(r.text, "html.parser")

    products = set()

    # 🔥 ВАЖНЫЙ jSHOP вариант
    blocks = soup.select(".jshop_product, .product, .product-item")

    print("BLOCKS:", len(blocks))

    for b in blocks:
        a = b.find("a", href=True)
        if a:
            products.add(urljoin(BASE, a["href"]))

    # 🔥 fallback (иногда jShop кладёт иначе)
    if not products:
        for a in soup.select("a[href*='product']"):
            products.add(urljoin(BASE, a["href"]))

    return list(products)


# =========================
# PRODUCT PARSE
# =========================
def parse_product(url):

    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    title = soup.select_one(".ttl.md.mb25")
    title = title.text.strip() if title else "-"

    art = soup.select_one(".prod-ean")
    art = art.text.replace("Артикул:", "").strip() if art else "-"

    # 🔥 ВАЖНО: правильная цена (INS + block_price)
    price = soup.select_one("#block_price, .prod_price, ins .woocommerce-Price-amount")
    price = price.text.strip() if price else "-"

    stock = soup.select_one(".avail, .prod-not-avail")
    stock = stock.text.strip() if stock else "-"

    return art, title, price, stock, url


# =========================
# RUN
# =========================
def run_parser():

    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True)

    try:
        login()

        cat = get_category()
        print("CATEGORY:", cat)

        wb = Workbook()
        ws = wb.active
        ws.append(["Артикул", "Название", "Цена", "Статус", "URL"])

        products = get_products(cat)

        print("PRODUCTS:", len(products))

        total = len(products)
        seen = set()

        for i, p in enumerate(products, 1):

            if p in seen:
                continue
            seen.add(p)

            art, title, price, stock, url = parse_product(p)

            if title == "-" or price == "-":
                continue

            ws.append([art, title, price, stock, url])

            percent = int(i / total * 100) if total else 100
            update_progress(percent)

            print(f"[{percent}%] {title} | {price}")

            time.sleep(0.3)

        wb.save(FILE_PATH)

        update_progress(100)
        print("DONE")

    finally:
        set_status(False)
        set_lock(False)


if __name__ == "__main__":
    run_parser()
