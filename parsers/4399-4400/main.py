import os
import time
import json
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from urllib.parse import urljoin

BASE = "https://jumpex.com.ua"

LOGIN_URL = BASE + "/user/loginsave"

LOGIN = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

OUTPUT_DIR = os.path.abspath("output/4399-4400")

FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)


# =========================
# STATUS
# =========================
def set_status(running: bool):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": 0,
            "time": datetime.now().strftime("%d.%m %H:%M")
        }, f, ensure_ascii=False, indent=2)


def update_progress(percent):
    if not os.path.exists(STATUS_PATH):
        return

    try:
        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = percent

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass


def set_lock(state: bool):
    if state:
        with open(LOCK_FILE, "w") as f:
            f.write("running")
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


def is_locked():
    return os.path.exists(LOCK_FILE)


# =========================
# SESSION
# =========================
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0"
})


# =========================
# LOGIN (requests)
# =========================
def login():
    """
    login form:
    username
    passwd
    """
    r = session.get(BASE + "/login")
    soup = BeautifulSoup(r.text, "html.parser")

    token_input = soup.find("input", {"type": "hidden"})

    data = {
        "username": LOGIN,
        "passwd": PASSWORD,
    }

    # если есть hidden token
    if token_input:
        data[token_input.get("name")] = token_input.get("value", "1")

    r = session.post(LOGIN_URL, data=data)

    print("LOGIN STATUS:", r.status_code)


# =========================
# CATEGORY (TEST :1)
# =========================
def get_category():
    url = BASE + "/"

    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    # берем первую основную категорию
    a = soup.select_one("li.nav-item.parent a[href]")

    if not a:
        raise Exception("NO CATEGORY FOUND")

    return urljoin(BASE, a["href"])


# =========================
# LOAD MORE
# =========================
def load_all_products(cat_url):
    products = set()

    page = 1

    while True:
        url = cat_url + f"?page={page}"

        r = session.get(url)
        soup = BeautifulSoup(r.text, "html.parser")

        items = soup.select(".product a[href]")

        if not items:
            break

        for i in items:
            href = i.get("href")
            if href and "/product" in href:
                products.add(urljoin(BASE, href))

        print(f"PAGE {page} OK -> {len(items)}")

        page += 1

        if page > 50:
            break

    return list(products)


# =========================
# PRODUCT PARSE
# =========================
def parse_product(url):
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    # TITLE
    title = soup.select_one(".ttl.md.mb25")
    title = title.text.strip() if title else "-"

    # ART
    art = soup.select_one(".prod-ean")
    art = art.text.strip().replace("Артикул:", "").strip() if art else "-"

    # PRICE (ВАЖНО: ТОЛЬКО ins = текущая цена)
    price = soup.select_one("ins .woocommerce-Price-amount, .prod_price")
    price = price.text.strip() if price else "-"

    # STOCK
    stock = soup.select_one(".avail, .prod-not-avail")
    stock = stock.text.strip() if stock else "-"

    return art, title, price, stock, url


# =========================
# MAIN
# =========================
def run_parser():

    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True)

    try:
        login()

        # 🔥 ТЕСТ :1 (одна категория)
        cat = get_category()

        print("CATEGORY:", cat)

        wb = Workbook()
        ws = wb.active

        ws.append(["Артикул", "Название", "Цена", "Статус", "Ссылка"])

        products = load_all_products(cat)

        total = len(products)
        print("PRODUCTS:", total)

        seen = set()

        for i, p in enumerate(products, 1):

            if p in seen:
                continue

            seen.add(p)

            art, title, price, stock, url = parse_product(p)

            if title == "-" or price == "-":
                continue

            ws.append([art, title, price, stock, url])

            percent = int(i / total * 100)
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
