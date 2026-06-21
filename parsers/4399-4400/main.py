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


# =========================
# OUTPUT
# =========================

OUTPUT_DIR = os.path.abspath("output/4399-4400")

FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_4399-4400_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


# =========================
# STATUS
# =========================

def set_status(running=True, progress=0, done=0, total=0):

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": progress,
            "done": done,
            "total": total,
            "time": datetime.now().strftime("%H:%M:%S")
        }, f, ensure_ascii=False, indent=2)


# =========================
# LOCK (СТАБИЛЬНЫЙ)
# =========================

def is_locked():
    if not os.path.exists(LOCK_FILE):
        return False

    age = time.time() - os.path.getmtime(LOCK_FILE)

    if age > 3600:
        os.remove(LOCK_FILE)
        return False

    return True


def set_lock(state: bool):
    if state:
        with open(LOCK_FILE, "w") as f:
            f.write(str(time.time()))
    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


# =========================
# DRIVER (ВАЖНО: НЕ ГЛОБАЛЬНЫЙ)
# =========================

def get_driver():

    options = Options()
    options.add_argument("--headless=new")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    driver = webdriver.Chrome(options=options)
    driver.set_page_load_timeout(30)

    return driver


# =========================
# LOGIN
# =========================

def login(driver):

    driver.get(BASE + "/login")
    time.sleep(3)

    driver.find_element(By.ID, "jlusername").send_keys(LOGIN)
    driver.find_element(By.ID, "jlpassword").send_keys(PASSWORD)
    driver.find_element(By.ID, "jlpassword").send_keys(Keys.RETURN)

    time.sleep(5)


# =========================
# CATEGORIES
# =========================

def get_categories(driver):

    r = requests.get(BASE + "/", headers={"User-Agent": "Mozilla/5.0"}, timeout=30)
    soup = BeautifulSoup(r.text, "html.parser")

    cats = set()

    for li in soup.select("li.nav-item.parent"):
        a = li.find("a", href=True)
        if not a:
            continue

        href = a.get("href")

        if href and href.startswith("/"):
            parts = href.strip("/").split("/")
            if len(parts) == 1:
                cats.add(BASE + href)

    return list(cats)


# =========================
# PRODUCTS
# =========================

def get_products_from_category(driver, cat_url):

    driver.get(cat_url)
    time.sleep(3)

    links = set()

    cards = driver.find_elements(By.CSS_SELECTOR, ".product, .product-item, .jshop_product")

    for c in cards:
        try:
            a = c.find_element(By.CSS_SELECTOR, "a[href]")
            href = a.get_attribute("href")

            if not href:
                continue

            if "/product" not in href:
                continue

            links.add(href)

        except:
            continue

    return list(links)


# =========================
# PRODUCT PARSE
# =========================

def parse_product(driver, url):

    driver.get(url)
    time.sleep(2)

    try:
        title = driver.find_element(By.TAG_NAME, "h1").text.strip()
    except:
        return None

    try:
        sku = driver.find_element(By.CSS_SELECTOR, ".prod-ean").text.strip()
    except:
        sku = "-"

    try:
        price = driver.find_element(By.ID, "block_price").text.strip()
    except:
        price = "-"

    try:
        status = driver.find_element(By.CSS_SELECTOR, ".avail, .prod-not-avail").text.strip()
    except:
        status = "-"

    return {
        "title": title,
        "sku": sku,
        "price": price,
        "status": status,
        "url": url
    }


# =========================
# MAIN
# =========================

def run_parser():

    print("🚀 START PARSER")

    if is_locked():
        print("⛔ ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True, 0)

    driver = get_driver()

    try:

        login(driver)

        categories = get_categories(driver)

        print("📂 CATEGORIES:", len(categories))

        wb = Workbook()
        ws = wb.active

        ws.append(["Название", "SKU", "Цена", "Статус", "URL"])

        all_products = []

        # 🔥 TEMP TEST = 1 CATEGORY
        categories = categories[:1]

        for cat in categories:

            products = get_products_from_category(driver, cat)

            for p in products:
                all_products.append(p)

        total = len(all_products)

        print("📦 TOTAL:", total)

        done = 0

        for url in all_products:

            data = parse_product(driver, url)

            if not data:
                continue

            ws.append([
                data["title"],
                data["sku"],
                data["price"],
                data["status"],
                data["url"]
            ])

            done += 1
            progress = int(done / total * 100) if total else 100

            set_status(True, progress, done, total)

            if done % 10 == 0:
                wb.save(FILE_PATH)

            time.sleep(random.uniform(0.2, 0.4))

        wb.save(FILE_PATH)

        set_status(False, 100, done, total)

        print("✅ DONE")

    finally:
        set_lock(False)
        set_status(False, 100)
        driver.quit()


def main():
    run_parser()


if __name__ == "__main__":
    main()
