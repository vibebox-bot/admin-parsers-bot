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

# =========================
# TEST MODE :1
# =========================
ONLY_ONE_CATEGORY = True   # <-- ВОТ ЭТО ТВОЙ :1

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "LIVE.xlsx")


# =========================
# HELPERS
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
# LOGIN
# =========================

def login(page):
    print("LOGIN...")

    page.goto(BASE + "/index.php?route=account/login", wait_until="networkidle")

    page.fill('input[name="email"]', EMAIL)
    page.fill('input[name="password"]', PASSWORD)

    page.click('input[type="submit"], button[type="submit"]')

    page.wait_for_timeout(4000)

    print("LOGIN OK")


# =========================
# CATEGORIES
# =========================

def get_categories(page):
    page.goto(BASE, wait_until="networkidle")
    html = soup(page)

    cats = set()

    for a in html.select("a[href*='route=product/category']"):
        cats.add(full_url(a["href"]))

    cats = list(cats)

    # =========================
    # :1 MODE → ТОЛЬКО 1 КАТЕГОРИЯ
    # =========================
    if ONLY_ONE_CATEGORY:
        print("⚠ TEST MODE :1 → only first category")
        return cats[:1]

    return cats


# =========================
# LOAD ALL PRODUCTS FROM CATEGORY
# =========================

def crawl_category(page, cat):
    print("CATEGORY:", cat)

    products = set()

    for i in range(1, 200):

        url = f"{cat}&page={i}" if "?" in cat else f"{cat}?page={i}"

        page.goto(url, wait_until="networkidle")
        html = soup(page)

        links = set()

        for a in html.find_all("a", href=True):
            if "product_id" in a["href"]:
                links.add(full_url(a["href"]))

        print(f"PAGE {i} OK -> {len(links)}")

        if not links and i > 1:
            break

        products.update(links)
        time.sleep(0.3)

    return list(products)


# =========================
# PRODUCT PARSE
# =========================

def parse_product(page, url):
    page.goto(url, wait_until="networkidle")
    html = soup(page)

    title = clean(html.select_one("h1").get_text()) if html.select_one("h1") else ""

    sku = clean(html.select_one(".prod-ean").get_text()) if html.select_one(".prod-ean") else ""

    price = clean(html.select_one(".prod_price").get_text()) if html.select_one(".prod_price") else ""

    status = ""
    if html.select_one(".avail"):
        status = clean(html.select_one(".avail").get_text())
    elif html.select_one(".prod-not-avail"):
        status = clean(html.select_one(".prod-not-avail").get_text())

    return [sku, title, price, status, url]


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
        with sync_playwright() as p:

            browser = p.chromium.launch(
                headless=True,
                args=[
                    "--no-sandbox",
                    "--disable-setuid-sandbox",
                    "--disable-dev-shm-usage"
                ]
            )

            page = browser.new_page()

            wb = Workbook()
            ws = wb.active
            ws.append(["SKU", "Title", "Price", "Status", "URL"])

            login(page)

            categories = get_categories(page)

            print("CATEGORIES:", len(categories))

            # =========================
            # 🔥 TEST MODE :1 CATEGORY
            # =========================
            categories = categories[:1]

            seen = set()
            total = 0

            for i, cat in enumerate(categories, 1):

                print("TEST CATEGORY:", cat)

                update_progress(0)

                products = crawl_category(page, cat)

                print("PRODUCTS FOUND:", len(products))

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

                        print(f"PRODUCT: {data[0]} | {data[2]}")

                    except Exception as e:
                        print("ERROR PRODUCT:", e)

                    time.sleep(0.2)

            os.makedirs(OUTPUT_DIR, exist_ok=True)
            wb.save(FILE_PATH)

            update_progress(100)

            print("DONE:", total)

            browser.close()

    finally:
        set_status(False)
        set_lock(False)

if __name__ == "__main__":
    run_parser()
