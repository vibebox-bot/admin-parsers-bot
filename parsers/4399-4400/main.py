import time
import random
import requests
import os
import json
from datetime import datetime

from bs4 import BeautifulSoup

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from openpyxl import Workbook

BASE = "https://jumpex.com.ua"
seen_categories = set()

LOGIN = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

# =========================
# MAIN.RU
# =========================

OUTPUT_DIR = os.path.abspath("output/4399-4400")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Харьковская_4399-4400_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "lock.txt"
)


def set_status(running: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "running": running,
                "progress": 0,
                "time": datetime.now().strftime("%d.%m %H:%M")
            },
            f,
            ensure_ascii=False,
            indent=2
        )


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
# DRIVER
# =========================
options = Options()
options.add_argument("--headless=new")
options.add_argument("--disable-gpu")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)
driver.set_page_load_timeout(30)

# =========================
# LOGIN
# =========================
def login():
    driver.get(BASE + "/login")
    time.sleep(3)

    driver.find_element(By.ID, "jlusername").send_keys(LOGIN)
    driver.find_element(By.ID, "jlpassword").send_keys(PASSWORD)
    driver.find_element(By.ID, "jlpassword").send_keys(Keys.RETURN)

    time.sleep(5)
    print("LOGIN OK")

# =========================
# ONLY MAIN CATEGORIES
# =========================
def get_categories():

    url = BASE + "/"

    r = requests.get(
        url,
        headers={"User-Agent": "Mozilla/5.0"},
        timeout=30
    )

    soup = BeautifulSoup(r.text, "html.parser")

    cats = set()

    for li in soup.select("li.nav-item.parent"):

        a = li.find("a", href=True)

        if not a:
            continue

        href = a.get("href")

        if not href:
            continue

        if not href.startswith("/"):
            continue

        # только категории первого уровня
        parts = [p for p in href.strip("/").split("/") if p]

        if len(parts) != 1:
            continue

        cats.add(BASE + href)

    return sorted(cats)
# =========================
# USD SWITCH
# =========================
def switch_to_usd():

    try:
        driver.get(
            BASE +
            "/ru/component/jshopping/?id_currency=2&back=/ru/"
        )

        time.sleep(2)

    except:
        pass

# =========================
# LOAD MORE
# =========================
def click_load_more():

    while True:

        try:
            buttons = driver.find_elements(By.CSS_SELECTOR, "button.autoScrollBtn")

            if not buttons:
                break

            btn = buttons[0]

            driver.execute_script("arguments[0].scrollIntoView(true);", btn)
            time.sleep(1)

            driver.execute_script("arguments[0].click();", btn)
            time.sleep(2)

        except:
            break

# =========================
# PRODUCT LINKS
# =========================
def get_products_from_category(cat_url):

    driver.get(cat_url)
    time.sleep(4)

    switch_to_usd()

    driver.get(cat_url)
    time.sleep(4)

    click_load_more()

    links = set()

    # берем ТОЛЬКО реальные товарные ссылки внутри карточек
    cards = driver.find_elements(By.CSS_SELECTOR, ".product, .product-item, .jshop_product")

    print("PRODUCT BLOCKS FOUND:", len(cards))

    for c in cards:
        try:
            a = c.find_element(By.CSS_SELECTOR, "a[href]")
            href = a.get_attribute("href")

            if not href:
                continue

            if BASE not in href:
                continue

            if href.rstrip("/") == cat_url.rstrip("/"):
                continue

            # ❌ мусор
            if any(x in href for x in [
                "/cart/",
                "/wishlist/",
                "/login",
                "#",
                "javascript",
                "?id_currency",
                "/filter",
                "/search",
                "/component/",
                "/tag",
                "/page"
            ]):
                continue

            # 💥 ГЛАВНЫЙ ФИЛЬТР: убираем категории и подкатегории
            slug = href.replace(BASE, "").strip("/")

            # категории НЕ имеют "-" как товар
            if "-" not in slug:
                continue

            links.add(href)

        except:
            continue

    return list(links)
# =========================
# PRODUCT
# =========================
def parse_product(url):

    driver.get(url)

    time.sleep(2.5)

    try:
        title = driver.find_element(By.TAG_NAME, "h1").text.strip()
    except:
        return None, None, None, None, url

    try:
        art = driver.find_element(By.CSS_SELECTOR, ".prod-ean").text.strip()
    except:
        art = "NO ART"

    try:
        price = driver.find_element(By.ID, "block_price").text.strip()
    except:
        price = "NO PRICE"

    # 🔥 пропускаем категории и мусор
    if art == "NO ART" and price == "NO PRICE":
        return None, None, None, None, url

    try:
        status = driver.find_element(
            By.CSS_SELECTOR,
            ".avail, .prod-not-avail"
        ).text.strip()
    except:
        status = "NO STATUS"

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

        login()

        cats = get_categories()

        print("CATEGORIES:", len(cats))

        for c in cats:
            print(c)

        wb = Workbook()
        ws = wb.active

        ws.append([
            "Артикул",
            "Название",
            "Цена",
            "Статус",
            "Ссылка"
        ])

        seen = set()
        total = 0

        for i, cat in enumerate(cats, 1):

            percent = int(i / len(cats) * 100)
            update_progress(percent)

            print(f"\n[CAT {i}/{len(cats)}]")
            print("CATEGORY:", cat)

            try:
                products = get_products_from_category(cat)
            except Exception as e:
                print("CATEGORY ERROR:", e)
                continue

            print("PRODUCTS:", len(products))
            print("SAMPLE:", list(products)[:5])

            for p in products:

                if p in seen:
                    continue

                seen.add(p)

                try:

                    art, title, price, status, url = parse_product(p)

                    if not title:
                        continue

                    if not title or not price:
                        continue

                    ws.append([
                        art,
                        title,
                        price,
                        status,
                        url
                    ])

                    total += 1

                    print(
                        total,
                        art,
                        price,
                        status
                    )

                except:
                    pass

                time.sleep(
                    random.uniform(0.2, 0.5)
                )

        os.makedirs(OUTPUT_DIR, exist_ok=True)
        wb.save(FILE_PATH)

        update_progress(100)

        driver.quit()

        print("\nDONE")
        print("TOTAL:", total)

    finally:
        set_status(False)
        set_lock(False)


if __name__ == "__main__":
    run_parser()