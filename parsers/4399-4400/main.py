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

LOGIN = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(OUTPUT_DIR, "TEST_JUMPEX.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

TEST_CATEGORY_LIMIT = 1  # 🔥 ВАЖНО: пока 1 категория


# =========================
# STATUS
# =========================
def update_progress(p):
    try:
        if not os.path.exists(STATUS_PATH):
            return

        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = p

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except:
        pass


# =========================
# DRIVER
# =========================
options = Options()
options.add_argument("--headless=new")
options.add_argument("--no-sandbox")
options.add_argument("--disable-dev-shm-usage")

driver = webdriver.Chrome(options=options)


# =========================
# LOGIN
# =========================
def login():
    driver.get(BASE + "/login")
    time.sleep(3)

    driver.find_element(By.ID, "jlusername").send_keys(LOGIN)
    driver.find_element(By.ID, "jlpassword").send_keys(PASSWORD)
    driver.find_element(By.ID, "jlpassword").send_keys(Keys.RETURN)

    time.sleep(4)
    print("LOGIN OK")


# =========================
# CATEGORIES (ONLY MAIN)
# =========================
def get_categories():
    r = requests.get(BASE, headers={"User-Agent": "Mozilla/5.0"})
    soup = BeautifulSoup(r.text, "html.parser")

    cats = []

    for a in soup.select("li.nav-item.parent > a"):
        href = a.get("href")

        if not href:
            continue

        if href.startswith("/"):
            cats.append(BASE + href)

    return cats


# =========================
# LOAD MORE BUTTON
# =========================
def load_all():
    while True:
        try:
            btn = driver.find_elements(By.CSS_SELECTOR, "button.autoScrollBtn")
            if not btn:
                break

            driver.execute_script("arguments[0].click();", btn[0])
            time.sleep(2)
        except:
            break


# =========================
# PRODUCTS
# =========================
def get_products(cat_url):
    driver.get(cat_url)
    time.sleep(3)

    load_all()

    links = set()

    cards = driver.find_elements(By.CSS_SELECTOR, ".product, .product-item, .jshop_product")

    for c in cards:
        try:
            a = c.find_element(By.CSS_SELECTOR, "a[href]")
            href = a.get_attribute("href")

            if BASE in href:
                links.add(href)

        except:
            continue

    return list(links)


# =========================
# PARSE PRODUCT
# =========================
def parse_product(url):
    driver.get(url)
    time.sleep(2)

    try:
        title = driver.find_element(By.CSS_SELECTOR, ".ttl").text.strip()
    except:
        title = "-"

    try:
        art = driver.find_element(By.CSS_SELECTOR, ".prod-ean").text.strip()
    except:
        art = "-"

    # ❌ ЦЕНА ЗАКРЫТА БЕЗ ЛОГИНА
    try:
        price = driver.find_element(By.CSS_SELECTOR, ".prod_price").text.strip()
    except:
        price = "HIDDEN"

    try:
        status = driver.find_element(By.CSS_SELECTOR, ".avail, .prod-not-avail").text.strip()
    except:
        status = "-"

    return art, title, price, status, url


# =========================
# RUN
# =========================
def run_parser():

    print("🚀 STARTED TEST PARSER")

    login()

    cats = get_categories()[:TEST_CATEGORY_LIMIT]  # 🔥 только 1 категория

    wb = Workbook()
    ws = wb.active

    ws.append(["Артикул", "Название", "Цена", "Статус", "URL"])

    total = 0

    for i, cat in enumerate(cats, 1):

        print("CATEGORY:", cat)

        products = get_products(cat)

        print("FOUND:", len(products))

        for p in products:

            art, title, price, status, url = parse_product(p)

            ws.append([art, title, price, status, url])

            total += 1

            print(total, title, price)

            update_progress(int((i / len(cats)) * 100))

            time.sleep(random.uniform(0.2, 0.5))

    wb.save(FILE_PATH)

    driver.quit()

    print("DONE TOTAL:", total)


if __name__ == "__main__":
    run_parser()
