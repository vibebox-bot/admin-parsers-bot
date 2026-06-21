import requests
import time
import os
import json
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

BASE = "https://jumpex.com.ua"

LOGIN = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

# =========================
# MODE
# =========================
TEST_MODE = 1   # 🔥 1 = только 1 категория, 0 = все

# =========================
# OUTPUT
# =========================
OUTPUT_DIR = os.path.abspath("output/jumpex")

FILE_PATH = os.path.join(OUTPUT_DIR, "jumpex_live.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)

session = requests.Session()

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# =========================
# STATUS
# =========================
def set_status(progress=0, running=True, text=""):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": progress,
            "text": text,
            "time": datetime.now().strftime("%H:%M:%S")
        }, f, ensure_ascii=False, indent=2)


# =========================
# LOCK
# =========================
def set_lock(state):
    if state:
        with open(LOCK_FILE, "w") as f:
            f.write("1")
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


def is_locked():
    return os.path.exists(LOCK_FILE)


# =========================
# LOGIN
# =========================
def login():
    r = session.get(BASE + "/login", headers=HEADERS)
    soup = BeautifulSoup(r.text, "html.parser")

    form = soup.find("form")

    data = {}

    for i in form.find_all("input"):
        name = i.get("name")
        if name:
            data[name] = i.get("value", "")

    data["username"] = LOGIN
    data["passwd"] = PASSWORD

    action = form.get("action")

    session.post(BASE + action, data=data, headers=HEADERS)

    print("LOGIN OK")


# =========================
# CATEGORIES
# =========================
def get_categories():
    r = session.get(BASE + "/", headers=HEADERS)
    soup = BeautifulSoup(r.text, "html.parser")

    cats = []

    for a in soup.select("li.nav-item.parent > a[href]"):
        href = a.get("href")

        if href and href.startswith("/"):
            cats.append(BASE + href)

    return cats


# =========================
# PRODUCTS LINKS
# =========================
def get_products(cat_url):
    page = 1
    links = set()

    while True:
        url = f"{cat_url}?page={page}"
        r = session.get(url, headers=HEADERS)

        if r.status_code != 200:
            break

        soup = BeautifulSoup(r.text, "html.parser")

        items = soup.select(".product a[href]")

        if not items:
            break

        for i in items:
            href = i.get("href")
            if href and "/product" in href:
                links.add(BASE + href)

        print(f"PAGE {page} | {len(items)} items")

        page += 1
        time.sleep(0.3)

    return list(links)


# =========================
# PRODUCT PARSE
# =========================
def parse_product(url):
    r = session.get(url, headers=HEADERS)
    soup = BeautifulSoup(r.text, "html.parser")

    title = soup.select_one(".ttl.md.mb25")
    title = title.text.strip() if title else "-"

    art = soup.select_one(".prod-ean")
    art = art.text.replace("Артикул:", "").strip() if art else "-"

    price = soup.select_one(".prod_price")
    price = price.text.strip() if price else "-"

    avail = soup.select_one(".avail")
    not_avail = soup.select_one(".prod-not-avail")

    if avail:
        status = avail.text.strip()
    elif not_avail:
        status = not_avail.text.strip()
    else:
        status = "-"

    return art, title, price, status, url


# =========================
# RUN
# =========================
def run():

    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(0, True, "START")

    try:
        login()

        cats = get_categories()

        # 🔥 TEST MODE (1 категория)
        if TEST_MODE == 1:
            cats = cats[:1]

        wb = Workbook()
        ws = wb.active

        ws.append(["Артикул", "Название", "Цена", "Наличие", "URL"])

        total_cats = len(cats)
        done_products = 0

        for i, cat in enumerate(cats, 1):

            cat_progress = int(i / total_cats * 100)

            set_status(cat_progress, True, f"CAT {i}/{total_cats}")

            print("\n====================")
            print("CATEGORY:", cat)

            products = get_products(cat)

            total_products = len(products)

            for j, p in enumerate(products, 1):

                art, title, price, status, url = parse_product(p)

                ws.append([art, title, price, status, url])

                done_products += 1

                progress = int((j / total_products) * 100) if total_products else 0

                print(done_products, title, price)

                set_status(progress, True, title)

                time.sleep(0.2)

        wb.save(FILE_PATH)

        set_status(100, False, "DONE")
        print("DONE")

    finally:
        set_lock(False)


if __name__ == "__main__":
    run()
