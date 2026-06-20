import os
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

from openpyxl import Workbook

# =========================
# CONFIG
# =========================

BASE_URL = "https://melad.com.ua"
LOGIN_URL = "https://melad.com.ua/login/"

EMAIL = "titorangelina@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/Melad")
FILE_PATH = os.path.join(OUTPUT_DIR, "Melad_ALL.xlsx")

import json
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")


def update_progress(percent):
    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        if not os.path.exists(STATUS_PATH):
            return

        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = int(percent)

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    except:
        pass

# =========================
# DRIVER
# =========================

driver = None
wait = None


# =========================
# LOGIN
# =========================

def login():
    print("LOGIN START")

    driver.get(LOGIN_URL)
    time.sleep(2)

    driver.find_element(By.NAME, "email").send_keys(EMAIL)
    driver.find_element(By.NAME, "password").send_keys(PASSWORD)
    driver.find_element(By.CSS_SELECTOR, "input[type='submit']").click()

    time.sleep(3)

    print("LOGIN DONE")


# =========================
# GET CATEGORIES
# =========================

def get_categories():
    print("LOADING HOME FOR CATEGORIES...")

    driver.get(BASE_URL)
    time.sleep(3)

    categories = []
    seen = set()

    # 🔥 берем только основные категории
    blocks = driver.find_elements(By.CSS_SELECTOR, ".has-children")

    for block in blocks:
        try:
            a = block.find_element(By.CSS_SELECTOR, "a")
            href = a.get_attribute("href")

            if not href:
                continue

            if "melad.com.ua" not in href:
                continue

            if any(x in href for x in ["login", "account", "cart", "checkout"]):
                continue

            if href in seen:
                continue

            seen.add(href)
            categories.append(href)

        except:
            continue

    print("CATEGORIES FOUND:", len(categories))
    for c in categories:
        print("CAT:", c)

    return categories


# =========================
# GET PRODUCTS ON PAGE
# =========================

def get_products():
    time.sleep(2)

    cards = driver.find_elements(By.CSS_SELECTOR, ".product-thumb")

    results = []

    for c in cards:

        try:
            name = c.find_element(By.CSS_SELECTOR, ".caption a").text.strip()
        except:
            name = ""

        try:
            price = c.find_element(By.CSS_SELECTOR, ".price").text.strip()
        except:
            price = ""

        try:
            code = c.find_element(By.CSS_SELECTOR, ".kod_sku b").text.strip()
        except:
            code = ""

        try:
            status = c.find_element(By.CSS_SELECTOR, ".hidden-sm").text.strip()
        except:
            try:
                status = c.find_element(By.CSS_SELECTOR, ".add_to_cart").text.strip()
            except:
                status = ""

        results.append((name, price, code, status))

    print("DEBUG CARDS FOUND:", len(cards))

    return results


# =========================
# CATEGORY PARSER (FULL PAGINATION)
# =========================

def parse_category(url, ws):
    print("OPEN CATEGORY:", url)
    print("\nCATEGORY:", url)

    page = 1
    seen = set()
    total = 0

    while True:

        current_url = url if page == 1 else f"{url}?page={page}"

        print("PAGE:", page, current_url)

        driver.get(current_url)
        time.sleep(3)

        products = get_products()

        print("FOUND:", len(products))

        if not products:
            print("NO PRODUCTS -> STOP CATEGORY")
            break

        new_items = 0

        for p in products:

            key = (p[0], p[2])

            if key in seen:
                continue

            seen.add(key)
            new_items += 1
            total += 1

            ws.append([
                url,
                p[0],
                p[1],
                p[2],
                p[3]
            ])

        print("NEW:", new_items)

        if new_items == 0:
            print("NO NEW ITEMS -> STOP")
            break

        page += 1

        if page > 50:
            print("SAFETY STOP")
            break

    print("TOTAL IN CATEGORY:", total)


# =========================
# RUN
# =========================

def run_parser():

    global driver, wait

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    options = Options()
    options.add_argument("--start-maximized")

    driver = webdriver.Chrome(options=options)
    wait = WebDriverWait(driver, 20)

    wb = Workbook()
    ws = wb.active
    ws.title = "melad"

    ws.append(["Category", "Name", "Price", "Code", "Status"])

    try:
        login()

        categories = get_categories()
        total_cats = len(categories)

        print("DEBUG CATEGORIES COUNT:", len(categories))
        print(categories[:3])

        for i, cat in enumerate(categories, 1):

            percent = int((i / total_cats) * 100)
            update_progress(percent)

            print(f"CAT {i}/{total_cats}")

            parse_category(cat, ws)

        wb.save(FILE_PATH)

        print("\nDONE:", FILE_PATH)

    finally:
        try:
            driver.quit()
        except:
            pass


# IMPORTANT FOR IMPORT
run_parser = run_parser


if __name__ == "__main__":
    run_parser()