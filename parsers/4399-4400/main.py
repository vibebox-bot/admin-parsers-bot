import os
import time
from playwright.sync_api import sync_playwright

# =========================
# ⚙️ CONFIG
# =========================

ONLY_ONE_CATEGORY = True  # 👈 :1 режим (потом False уберём)

BASE_URL = "https://jumpex.com.ua"

LOGIN_URL = BASE_URL + "/login"
CATALOG_URL = BASE_URL + "/instrumenty-i-oborudovanie"


# =========================
# 🔥 FIX PLAYWRIGHT (ВАЖНО)
# =========================
os.system("python -m playwright install --with-deps chromium")


# =========================
# LOGIN
# =========================
def login(page):
    page.goto(LOGIN_URL)

    page.fill("#jlusername", "angelinatitor@gmail.com")
    page.fill("#jlpassword", "380931937922")

    page.click("button[type=submit]")

    page.wait_for_timeout(3000)
    print("LOGIN: OK")


# =========================
# LOAD FULL CATEGORY (AUTO SCROLL)
# =========================
def load_full_category_html(page, url):
    page.goto(url)

    print(f"CATEGORY: {url}")

    last_height = 0

    while True:
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)

        try:
            page.click("button.autoScrollBtn", timeout=2000)
            print("CLICK: Показати ще")
            time.sleep(2)
        except:
            pass

        new_height = page.evaluate("document.body.scrollHeight")

        if new_height == last_height:
            break

        last_height = new_height

    return page.content()


# =========================
# PARSE PRODUCTS
# =========================
def parse_products(page):
    products = page.query_selector_all(".product")

    result = []

    for p in products:
        try:
            title = p.query_selector(".ttl.md.mb25")
            price = p.query_selector(".prod_price")
            sku = p.query_selector(".prod-ean.mb60")
            avail = p.query_selector(".avail, .prod-not-avail")

            data = {
                "title": title.inner_text().strip() if title else "",
                "price": price.inner_text().strip() if price else "",
                "sku": sku.inner_text().strip() if sku else "",
                "avail": avail.inner_text().strip() if avail else "",
            }

            result.append(data)

        except:
            continue

    return result


# =========================
# MAIN PARSER
# =========================
def run_parser():
    with sync_playwright() as p:

        browser = p.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox",
                "--disable-setuid-sandbox"
            ]
        )

        page = browser.new_page()

        login(page)

        categories = [
            CATALOG_URL
        ]

        # 👇 :1 режим
        if ONLY_ONE_CATEGORY:
            categories = categories[:1]

        all_products = []

        for cat in categories:
            html = load_full_category_html(page, cat)

            page.goto(cat)
            page.wait_for_timeout(2000)

            products = parse_products(page)

            print(f"PRODUCTS: {len(products)}")

            for i, pr in enumerate(products):
                percent = int((i + 1) / len(products) * 100)
                print(f"[{percent}%] {pr['title']} | {pr['price']}")

            all_products.extend(products)

        print("DONE")
        print("TOTAL:", len(all_products))

        browser.close()


# =========================
# ENTRY POINT
# =========================
if __name__ == "__main__":
    run_parser()
