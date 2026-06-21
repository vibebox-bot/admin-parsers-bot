import os
os.system("python -m playwright install --with-deps chromium")
import json
import time
import random
from datetime import datetime

import requests
from bs4 import BeautifulSoup

from playwright.sync_api import sync_playwright
from openpyxl import Workbook

# =========================
# CONFIG
# =========================

BASE = "https://jumpex.com.ua"

LOGIN = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

OUTPUT_DIR = os.path.abspath("output/4399-4400")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Харьковская_4399-4400_LIVE.xlsx"
)

STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


# =========================
# STATUS
# =========================

def set_status(running: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

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
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if state:
        with open(LOCK_FILE, "w") as f:
            f.write("running")
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


def is_locked():
    return os.path.exists(LOCK_FILE)


# =========================
# LOGIN (requests)
# =========================

def login_session():
    s = requests.Session()

    s.headers.update({
        "User-Agent": "Mozilla/5.0"
    })

    s.get(BASE + "/login")

    payload = {
        "username": LOGIN,
        "passwd": PASSWORD
    }

    s.post(BASE + "/user/loginsave", data=payload)

    print("LOGIN: OK")
    return s


# =========================
# CATEGORIES (1 LEVEL ONLY)
# =========================

def get_categories(session):
    r = session.get(BASE + "/")
    soup = BeautifulSoup(r.text, "html.parser")

    cats = []

    for a in soup.select("li.nav-item.parent > a"):
        href = a.get("href")

        if not href or not href.startswith("/"):
            continue

        parts = href.strip("/").split("/")

        # только главные категории
        if len(parts) != 1:
            continue

        cats.append(BASE + href)

    return cats


# =========================
# PLAYWRIGHT LOAD MORE
# =========================

def load_full_category_html(url):
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()

        page.goto(url, timeout=60000)
        time.sleep(2)

        while True:
            try:
                btn = page.query_selector("button.autoScrollBtn")
                if not btn:
                    break

                btn.click()
                time.sleep(1)

            except:
                break

        html = page.content()
        browser.close()

        return html


# =========================
# PARSE PRODUCTS FROM LIST
# =========================

def parse_category_products(html):
    soup = BeautifulSoup(html, "html.parser")

    products = []

    blocks = soup.select(".product, .product-item, .jshop_product")

    print("BLOCKS:", len(blocks))

    for b in blocks:
        a = b.find("a", href=True)

        if not a:
            continue

        href = a["href"]

        if href.startswith("/"):
            href = BASE + href

        products.append(href)

    return list(set(products))


# =========================
# PARSE PRODUCT PAGE
# =========================

def parse_product(session, url):
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    try:
        title = soup.select_one("h1.ttl.md.mb25").get_text(strip=True)
    except:
        return None

    try:
        art = soup.select_one(".prod-ean").get_text(strip=True).replace("Артикул:", "").strip()
    except:
        art = "NO ART"

    try:
        price = soup.select_one(".prod_price").get_text(strip=True)
    except:
        price = "NO PRICE"

    status_el = soup.select_one(".avail, .prod-not-avail")

    status = status_el.get_text(strip=True) if status_el else "NO STATUS"

    return art, title, price, status, url


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
        session = login_session()

        cats = get_categories(session)

        print("CATEGORIES:", len(cats))

        # =========================
        # TEST MODE :1
        # =========================
        cats = cats[:1]

        wb = Workbook()
        ws = wb.active

        ws.append(["Артикул", "Название", "Цена", "Статус", "Ссылка"])

        seen = set()
        total = 0

        for i, cat in enumerate(cats, 1):

            percent = int(i / len(cats) * 100)
            update_progress(percent)

            print(f"\nCATEGORY: {cat}")

            html = load_full_category_html(cat)
            products = parse_category_products(html)

            print("PRODUCTS:", len(products))

            for p in products:

                if p in seen:
                    continue

                seen.add(p)

                data = parse_product(session, p)

                if not data:
                    continue

                art, title, price, status, url = data

                ws.append([art, title, price, status, url])

                total += 1

                print(f"[{percent}%] {title} | {price}")

                time.sleep(random.uniform(0.1, 0.3))

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        wb.save(FILE_PATH)

        update_progress(100)

        print("\nDONE:", total)

    finally:
        set_status(False)
        set_lock(False)


if __name__ == "__main__":
    run_parser()
