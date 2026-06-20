import os
import json
import re
import time
from datetime import datetime

from playwright.sync_api import sync_playwright
from bs4 import BeautifulSoup
from openpyxl import Workbook

# =========================
# CONFIG
# =========================

BASE = "https://www.jmaxtvshop.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


# =========================
# STATUS / LOCK
# =========================

def set_status(running: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": 0,
            "time": datetime.now().strftime("%d.%m %H:%M")
        }, f, ensure_ascii=False, indent=2)


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

        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = percent

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )

    except:
        pass

# =========================
# HELPERS (ТВОЙ КОД)
# =========================

def clean(x):
    return re.sub(r"\s+", " ", x).strip() if x else ""


def soup(page):
    return BeautifulSoup(page.content(), "html.parser")


def full_url(href):
    if href.startswith("http"):
        return href
    return BASE + "/" + href.lstrip("/")


# =========================
# LOGIN (НЕ ТРОГАЮ)
# =========================

def login(page):
    print("LOGIN...")

    page.goto(BASE + "/index.php?route=account/login", wait_until="networkidle")

    page.fill('input[name="email"]', EMAIL)
    page.fill('input[name="password"]', PASSWORD)

    page.click('input[type="submit"], button[type="submit"]')

    page.wait_for_timeout(5000)

    if "logout" in page.content().lower():
        print("LOGIN OK")
    else:
        print("LOGIN CHECK")


# =========================
# CATEGORIES
# =========================

def get_categories(page):
    page.goto(BASE, wait_until="networkidle")
    html = soup(page)

    cats = set()

    for a in html.select(".menu-wrapper a[href]"):
        href = a.get("href")
        if href and "javascript" not in href:
            cats.add(full_url(href))

    for a in html.select("a[href*='route=product/category']"):
        cats.add(full_url(a["href"]))

    return list(cats)


# =========================
# PRODUCTS
# =========================

def crawl_category(page, cat):
    print("CATEGORY:", cat)

    all_links = set()

    for i in range(1, 200):

        url = f"{cat}&page={i}" if "?" in cat else f"{cat}?page={i}"

        page.goto(url, wait_until="networkidle")
        html = soup(page)

        links = set()

        for a in html.find_all("a", href=True):
            href = a["href"]
            if "product_id" in href:
                links.add(full_url(href))

        print(f"PAGE {i}: {len(links)}")

        if len(links) == 0 and i > 1:
            break

        all_links.update(links)
        time.sleep(0.5)

    return list(all_links)


def parse_product(page, url):
    page.goto(url, wait_until="networkidle")
    html = soup(page)

    title = clean(html.select_one("h1").get_text()) if html.select_one("h1") else ""
    price = clean(html.select_one(".product-page__price.price").get_text()) if html.select_one(".product-page__price.price") else ""
    sku = clean(html.select_one(".product-data__item.model").get_text()) if html.select_one(".product-data__item.model") else ""
    status = clean(html.select_one(".product-page__cart.row-flex").get_text(" ")) if html.select_one(".product-page__cart.row-flex") else ""

    return [sku, title, price, status, url]


# =========================
# MAIN PARSER (SAFE WRAPPER)
# =========================

def run_parser():

    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True)

    try:
        with sync_playwright() as p:

            browser = p.chromium.launch(headless=False)
            page = browser.new_page()

            wb = Workbook()
            ws = wb.active
            ws.append(["SKU", "Title", "Price", "Status", "URL"])

            login(page)

            categories = get_categories(page)

            print("CATEGORIES:", len(categories))

            total_categories = len(categories)

            seen = set()
            total = 0

            for i, cat in enumerate(categories, 1):

                percent = int(i / total_categories * 100)
                update_progress(percent)

                products = crawl_category(page, cat)

                for url in products:

                    try:
                        data = parse_product(page, url)

                        if data[0] and data[0] in seen:
                            continue

                        if data[0]:
                            seen.add(data[0])

                        if not data[1]:
                            continue

                        ws.append(data)
                        total += 1

                        print("PRODUCT:", data[0])

                    except:
                        pass

                    time.sleep(0.2)

            os.makedirs(OUTPUT_DIR, exist_ok=True)
            wb.save(FILE_PATH)

            print("DONE:", total)

            update_progress(100)

            browser.close()

    finally:
        set_status(False)
        set_lock(False)


if __name__ == "__main__":
    run_parser()