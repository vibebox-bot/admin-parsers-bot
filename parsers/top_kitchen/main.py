import os
import json
import time
from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait

from openpyxl import Workbook

# =========================
# PATHS
# =========================

BASE_URL = "https://www.top-kitchen.com.ua"

OUTPUT_DIR = os.path.abspath("output/top_kitchen")
FILE_PATH = os.path.join(OUTPUT_DIR, "top_kitchen_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

# =========================
# STATUS
# =========================

def set_status(running: bool):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "running": running,
                "progress": 0,
                "time": datetime.now().strftime("%Y-%m-%d %H:%M")
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

        if not isinstance(data, dict):
            data = {}

        data["progress"] = percent

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
# CATEGORIES (НЕ ТРОГАЮ ЛОГИКУ)
# =========================

def get_categories():
    driver.get(BASE_URL)
    time.sleep(3)

    cats = []
    seen = set()

    main = driver.find_elements(By.CSS_SELECTOR, ".list-group__a")
    sub = driver.find_elements(By.CSS_SELECTOR, ".list-group__children-a")

    for el in main + sub:
        href = el.get_attribute("href")

        if not href:
            continue

        if "top-kitchen.com.ua" not in href:
            continue

        if href in seen:
            continue

        seen.add(href)
        cats.append(href)

    return cats

def get_product_links():
    time.sleep(2)

    cards = driver.find_elements(
        By.CSS_SELECTOR,
        ".product-thumb a[href], .product-layout a[href]"
    )

    links = []
    seen = set()

    for c in cards:
        href = c.get_attribute("href")

        if not href:
            continue

        if not href.endswith(".html"):
            continue

        if href in seen:
            continue

        seen.add(href)
        links.append(href)

    return links

def parse_product(url):
    driver.get(url)

    wait.until(
        lambda d: d.execute_script(
            "return document.readyState"
        ) == "complete"
    )

    time.sleep(1.5)

    try:
        name = driver.find_element(
            By.CSS_SELECTOR,
            ".heading-h1 h1"
        ).text.strip()
    except:
        name = ""

    try:
        model = driver.find_element(
            By.CSS_SELECTOR,
            ".product-data__item.model"
        ).text

        model = model.replace(
            "Код Товара:",
            ""
        ).strip()
    except:
        model = ""

    try:
        sku = driver.find_element(
            By.CSS_SELECTOR,
            ".product-data__item.sku"
        ).text

        sku = sku.replace(
            "Артикул:",
            ""
        ).strip()
    except:
        sku = ""

    try:
        price = driver.find_element(
            By.CSS_SELECTOR,
            ".product-page__price.price"
        ).text.strip()
    except:
        price = ""

    qty = ""

    try:
        el = driver.find_element(
            By.CSS_SELECTOR,
            ".qty-indicator__bar"
        )

        qty = el.get_attribute(
            "data-original-title"
        ) or ""

    except:
        try:
            qty = driver.find_element(
                By.CSS_SELECTOR,
                ".qty-indicator"
            ).text.strip()
        except:
            qty = ""

    return name, model, sku, price, qty

def parse_category(url):
    seen = set()
    page = 1

    while True:

        current_url = (
            url
            if page == 1
            else f"{url}?page={page}"
        )

        driver.get(current_url)
        time.sleep(3)

        links = get_product_links()

        if not links:
            break

        new_links = []

        for link in links:
            if link not in seen:
                seen.add(link)
                new_links.append(link)

        if not new_links:
            break

        for link in new_links:
            try:
                name, model, sku, price, qty = parse_product(link)

                ws.append(
                    [
                        url,
                        name,
                        model,
                        sku,
                        price,
                        qty,
                        link
                    ]
                )

            except:
                pass

        page += 1

        if page > 50:
            break

# =========================
# MAIN
# =========================

def run_parser():
    global driver, wait, wb, ws

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # защита от дубля
    if is_locked():
        print("ALREADY RUNNING")
        return

    set_lock(True)
    set_status(True)

    try:
        options = Options()
        options.add_argument("--start-maximized")

        driver = webdriver.Chrome(options=options)
        wait = WebDriverWait(driver, 20)

        wb = Workbook()
        ws = wb.active
        ws.title = "top-kitchen"

        ws.append(
            [
                "Category",
                "Name",
                "Model",
                "SKU",
                "Price",
                "Availability",
                "URL"
            ]
        )

        categories = get_categories()  # [:1]
        total = len(categories)

        for i, cat in enumerate(categories, 1):

            percent = int(i / total * 100)
            update_progress(percent)

            parse_category(cat)

        update_progress(100)

        wb.save(FILE_PATH)

    finally:
        set_status(False)
        set_lock(False)

        try:
            driver.quit()
        except:
            pass

if __name__ == "__main__":
    run_parser()