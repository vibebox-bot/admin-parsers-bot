import os
import time
from playwright.sync_api import sync_playwright

# =========================
# ⚙️ CONFIG
# =========================

ONLY_ONE_CATEGORY = True  # 👈 :1 режим (потом уберём)

BASE_URL = "https://jumpex.com.ua"

LOGIN_URL = BASE_URL + "/login"
CATEGORY_URL = BASE_URL + "/instrumenty-i-oborudovanie"


# =========================
# 🔥 PLAYWRIGHT FIX (RAILWAY SAFE)
# =========================
def ensure_playwright():
    os.system("python -m playwright install --with-deps chromium")

ensure_playwright()


# =========================
# LOGIN
# =========================
def login(page):
    page.goto(LOGIN_URL)

    # ⚠️ поставь свои данные
    page.fill("#jlusername", "angelinatitor@gmail.com")
    page.fill("#jlpassword", "380931937922")

    page.click("button[type=submit]")

    page.wait_for_timeout(3000)
    print("LOGIN: OK")


# =========================
# LOAD FULL CATEGORY (AUTO SCROLL + BUTTON)
# =========================
def load_full_category(page, url):
    page.goto(url)
    print(f"CATEGORY: {url}")

    last_height = 0

    while True:
        # scroll вниз
        page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        time.sleep(2)

        # нажать "Показати ще"
        try:
            btn = page.query_selector("button.autoScrollBtn")
            if btn:
                btn.click()
                print("CLICK: Показати ще")
                time.sleep(2)
        except:
            pass

        new_height = page.evaluate("document.body.scrollHeight")

        if new_height == last_height:
            break

        last_height = new_height

    return page


# =========================
# PARSE PRODUCT CARD
# =========================
def parse_products(page):
    items = page.query_selector_all(".product, .product-item, .jshop_list_product")

    result = []

    for item in items:
        try:
            title = item.query_selector(".ttl, .product-title, h1")
            price = item.query_selector(".prod_price, .price")
            sku = item.query_selector(".prod-ean")
            avail = item.query_selector(".avail, .prod-not-avail")

            data = {
                "title": title.inner_text().strip() if title else "",
                "price": price.inner_text().strip() if price else "",
                "sku": sku.inner_text().strip() if sku else "",
                "avail": avail.inner_text().strip() if avail else "",
            }

            if data["title"]:
                result.append(data)

        except:
            continue

    return result


# =========================
# MAIN
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

        categories = [CATEGORY_URL]

        # 👇 :1 режим
        if ONLY_ONE_CATEGORY:
            categories = categories[:1]

        all_products = []

        for cat in categories:

            page = load_full_category(page, cat)

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
