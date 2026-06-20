import os
import json
import time
import threading
import sys

from datetime import datetime

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from openpyxl import Workbook

STOP_FLAG = False

# =========================
# PATHS
# =========================
BASE_DIR = os.path.abspath("output/4425-4426_Gold_Top")

FILE_PATH = os.path.join(BASE_DIR, "Харьковская_4425-4426_Gold_Top_LIVE.xlsx")
STATUS_PATH = os.path.join(BASE_DIR, "status.json")
LOCK_FILE = os.path.join(BASE_DIR, "lock.txt")


# =========================
# STATUS HELPERS
# =========================
def set_status(running: bool):
    os.makedirs(BASE_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({
            "running": running,
            "progress": 0,
            "time": datetime.now().strftime("%d.%m %H:%M")
        }, f, ensure_ascii=False, indent=2)


def set_lock(state: bool):
    os.makedirs(BASE_DIR, exist_ok=True)

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
# LOGIN DATA
# =========================
LOGIN_URL = "https://www.gold-tor.com.ua/index.php?route=account/login"
LOGIN = "Sawrun_05@icloud.com"
PASSWORD = "18022021"


# =========================
# DRIVER
# =========================
driver = None
wait = None


# =========================
# LOGIN (YOUR ORIGINAL - untouched logic)
# =========================
def login():
    driver.get(LOGIN_URL)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(2)

    driver.execute_script("""
        document.querySelectorAll('.modal-backdrop, .modal, .popup, .overlay')
        .forEach(e => e.remove());
    """)

    email = wait.until(
        EC.presence_of_all_elements_located((By.NAME, "email"))
    )
    email = next(e for e in email if e.is_displayed())

    password = wait.until(
        EC.presence_of_all_elements_located((By.NAME, "password"))
    )
    password = next(p for p in password if p.is_displayed())

    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", email)
    time.sleep(0.5)

    driver.execute_script("arguments[0].value='';", email)
    email.send_keys(LOGIN)

    driver.execute_script("arguments[0].value='';", password)
    password.send_keys(PASSWORD)
    password.send_keys(Keys.RETURN)

    time.sleep(4)
    print("LOGIN OK")


# =========================
# CATEGORIES
# =========================
def get_categories():
    driver.get("https://www.gold-tor.com.ua/")
    time.sleep(2)

    links = driver.find_elements(By.CSS_SELECTOR, "#d_category_menu_list a[href]")

    cats = []
    for l in links:
        href = l.get_attribute("href")
        if href and "gold-tor.com.ua" in href:
            if href not in cats:
                cats.append(href)

    return cats


# =========================
# SAFE TEXT
# =========================
def safe_text(css):
    try:
        return driver.find_element(By.CSS_SELECTOR, css).text.strip()
    except:
        return ""


# =========================
# PRODUCT PAGE
# =========================
def parse_product_page(url):
    driver.get(url)
    wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
    time.sleep(1.5)

    name = safe_text(".h2.my-4") or safe_text("h1")

    try:
        sku = driver.find_element(By.CSS_SELECTOR, ".mr-4.p-1.text-secondary").text.strip()
    except:
        sku = ""

    try:
        price = driver.find_element(By.CSS_SELECTOR, ".h2.m-0.text-nowrap").text.strip()
    except:
        price = ""

    try:
        status_el = driver.find_element(By.CSS_SELECTOR, ".alert")
        cls = status_el.get_attribute("class")

        if "alert-success" in cls:
            status = "В наличии"
        elif "alert-danger" in cls:
            status = "Нет в наличии"
        else:
            status = status_el.text.strip()
    except:
        status = "unknown"

    return name, sku, price, status


# =========================
# CATEGORY PARSER (FIXED SAFE LOOP)
# =========================
def parse_category(url):
    driver.get(url)
    time.sleep(2)

    seen = set()
    page = 1

    while True:
        if STOP_FLAG:
            return

        time.sleep(2)

        cards = driver.find_elements(By.CSS_SELECTOR, ".product-thumb, .product-layout, [class*='product']")

        new_links = []

        for c in cards:
            try:
                a = c.find_element(By.TAG_NAME, "a")
                href = a.get_attribute("href")

                if href and href not in seen:
                    seen.add(href)
                    new_links.append(href)
            except:
                pass

        if not new_links:
            break

        current_url = driver.current_url

        for link in new_links:
            if STOP_FLAG:
                return

            try:
                name, sku, price, status = parse_product_page(link)

                ws.append([
                    url,
                    name,
                    sku,
                    price,
                    status,
                    link
                ])

            except:
                pass

            driver.get(current_url)
            time.sleep(1)

        # pagination
        links = driver.find_elements(By.CSS_SELECTOR, ".pagination a")
        next_found = False

        for l in links:
            href = l.get_attribute("href")
            if href and f"page={page+1}" in href:
                driver.get(href)
                next_found = True
                break

        if not next_found:
            break

        page += 1


# =========================
# MAIN RUN (FOR BOT)
# =========================
def run_parser():
    global driver, wait, ws, wb

    os.makedirs(BASE_DIR, exist_ok=True)

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
        ws.title = "gold-tor"

        ws.append(["Category", "Name", "SKU", "Price", "Status", "URL"])

        login()

        categories = get_categories()   # [:1]

        total_categories = len(categories)

        for i, cat in enumerate(categories, 1):

            percent = int(i / total_categories * 100)
            update_progress(percent)
            time.sleep(0.1)

            parse_category(cat)


        wb.save(FILE_PATH)

        update_progress(100)



    finally:
        set_status(False)
        set_lock(False)

        try:
            driver.quit()
        except:
            pass
        os._exit(0)   # 💣 ВОТ ЭТО ДОБАВИТЬ В КОНЕЦ

if __name__ == "__main__":
    run_parser()