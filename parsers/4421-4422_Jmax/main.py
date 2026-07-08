import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import datetime

import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

# =========================
# CONFIG
# =========================
BASE = "http://www.jmaxtvshop.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": BASE
})

wb = None
ws = None


# =========================
# INIT
# =========================
def init():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 🧠 проверяем "живой ли процесс"
    if os.path.exists(LOCK_FILE):
        try:
            with open(LOCK_FILE, "r") as f:
                pid = int(f.read().strip())

            # проверяем существует ли процесс
            import psutil
            if psutil.pid_exists(pid):
                print("ALREADY RUNNING (live process)")
                exit()
            else:
                print("STALE LOCK → removing")
                os.remove(LOCK_FILE)

        except:
            print("BROKEN LOCK → removing")
            os.remove(LOCK_FILE)

    # создаём новый lock с PID
    with open(LOCK_FILE, "w") as f:
        f.write(str(os.getpid()))

def finish():
    global wb

    if wb:
        wb.save(FILE_PATH)

    if os.path.exists(LOCK_FILE):
        try:
            os.remove(LOCK_FILE)
        except:
            pass

# =========================
# STATUS
# =========================
def update_status(data):
    os.makedirs(os.path.dirname(STATUS_PATH), exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def set_status(**kwargs):
    # читаем старые данные
    data = {}

    if os.path.exists(STATUS_PATH):
        try:
            with open(STATUS_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
        except:
            data = {}

    # обновляем только то что пришло
    data.update(kwargs)

    if "user" not in data:
        data["user"] = USER

    if "file_path" not in data:
        data["file_path"] = FILE_PATH

    # всегда обновляем время только если не передали своё
    if "time" not in data:
        data["time"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    update_status(data)


def check_cancel():
    if os.path.exists(STATUS_PATH):
        try:
            with open(STATUS_PATH, "r", encoding="utf-8") as f:
                st = json.load(f)
            if st.get("canceled"):
                return True
        except:
            pass
    return False


# =========================
# HELPERS
# =========================
def get_soup(url):
    try:
        r = session.get(
            url,
            timeout=20,
            headers={"Connection": "close"}
        )
        return BeautifulSoup(r.text, "html.parser")
    except Exception as e:
        print("REQUEST ERROR:", url, e)
        return BeautifulSoup("", "html.parser")


def clean(text):
    return text.strip().replace("\n", " ").replace("\t", " ")

# =========================
# LOGIN
# =========================
def login():
    print("LOGIN...")

    login_url = BASE + "/index.php?route=account/login"

    # 1. сначала открываем страницу (ВАЖНО для cookies)
    r = session.get(login_url)

    soup = BeautifulSoup(r.text, "html.parser")

    # 2. берем action формы (иногда отличается)
    form = soup.select_one("form")
    if not form:
        print("LOGIN FAILED: no form")
        return False

    action = form.get("action")
    if not action:
        action = login_url

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    # 3. отправляем логин уже в action
    r2 = session.post(action, data=payload, allow_redirects=True)

    # DEBUG (очень важно)
    #print("DEBUG URL:", r2.url)
    #print("DEBUG STATUS:", r2.status_code)

    if "logout" in r2.text.lower() or "account/logout" in r2.text.lower():
        print("LOGIN OK")
        return True

    # доп проверка (OpenCart часто редиректит в account)
    if "account/account" in r2.url:
        print("LOGIN OK (redirect)")
        return True

    #print("LOGIN FAILED")
    #print("DEBUG TEXT:", r2.text[:500])
    return False
    
# =========================
# CATEGORIES
# =========================
def get_categories():
    soup = get_soup(BASE)

    cats = []
    for a in soup.select("#menu a"):
        href = a.get("href")
        if href and "route=product/category" in href:
            cats.append(href)

    cats = list(dict.fromkeys(cats))
    print("FOUND CATEGORIES:", len(cats))
    return cats


# =========================
# PRODUCTS LIST
# =========================
def load_products(cat_url):
    #print("CATEGORY:", cat_url)

    all_products = set()
    page = 1

    while True:
        if check_cancel():
            print("CANCELED")
            return []

        url = f"{cat_url}&page={page}"
        soup = get_soup(url)

        items = soup.select(".product-thumb.uni-item")

        #print(f"PAGE {page} -> ITEMS {len(items)}")

        if not items:
            break

        before = len(all_products)

        for item in items:
            a = item.select_one("a")
            if a:
                href = a.get("href")
                if href:
                    all_products.add(href)

        if len(all_products) == before:
            break

        page += 1
        time.sleep(0.3)

    return list(all_products)


# =========================
# PRODUCT PARSER
# =========================
def parse_product(url):
    soup = get_soup(url)

    title = ""
    sku = ""
    price = ""
    status = ""

    for i in range(3):
        try:
            soup = get_soup(url)
            break
        except:
            time.sleep(1)
    
    h1 = soup.select_one("h1")
    if h1:
        title = clean(h1.get_text())

    sku_tag = soup.select_one(".product-data__item.model")
    if sku_tag:
        sku = clean(sku_tag.get_text().replace("Код товара:", ""))

    price_tag = soup.select_one(".product-page__price")
    if price_tag:
        price = clean(price_tag.get_text())

    # =========================
    # RAW BUTTON TEXT (как на сайте)
    # =========================
    btn = soup.select_one("#button-cart span")
    
    if btn:
        status = clean(btn.get_text())
    else:
        status = "Нет кнопки"

    return [sku, title, price, status, url]


# =========================
# EXCEL
# =========================
def init_excel():
    global wb, ws

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "Title", "Price", "Status", "URL"])


    tmp = FILE_PATH + ".tmp"
    wb.save(tmp)
    os.replace(tmp, FILE_PATH)

# =========================
# RUN
# =========================
def run_parser():
    print("🚀 STARTED JMAX PARSER")
    init()
    seen = set()

    ok = login()
    if not ok:
        set_status(running=False, progress=0, error="LOGIN FAILED")
        return

    init_excel()

    set_status(
        running=True,
        canceled=False,
        progress=0,
        user=USER,
        file_path=FILE_PATH
    )

    categories = get_categories()

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

    total_categories = len(categories)
    total_products_global = 0

    # считаем заранее (чтобы был норм прогресс)
    all_products_map = {}

    for cat in categories:
        prods = load_products(cat)
        all_products_map[cat] = prods
        total_products_global += len(prods)

    done_products = 0

    for ci, cat in enumerate(categories):

        products = all_products_map[cat]

        for pi, p in enumerate(products):

            if check_cancel():
                set_status(running=False, canceled=True, progress=0)
                finish()
                return


            try:
                row = parse_product(p)
            
                sku = row[0]   # 👈 SKU это первый элемент
            
                if sku in seen:
                    continue
            
                seen.add(sku)
            
                ws.append(row)
            
            except Exception as e:
                print("ERROR:", p, e)
            

            done_products += 1

            progress = int((done_products / max(total_products_global, 1)) * 100)

            set_status(
                running=True,
                canceled=False,
                progress=progress,
                category_index=ci,
                product_index=pi,
                total_products=total_products_global,
                file_path=FILE_PATH
            )

            if done_products % 20 == 0 or progress == 100:
                print(f"PROGRESS: {progress}%")

            time.sleep(0.1)

    
    set_status(
        running=False,
        canceled=False,
        progress=100,
        user=USER,
        file_path=FILE_PATH
    )
    
    finish()

    print("DONE")
    print("✅ FINISHED JMAX PARSER")


if __name__ == "__main__":
    run_parser()
