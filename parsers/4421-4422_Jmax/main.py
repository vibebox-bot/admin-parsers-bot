import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Харьковская 4421-4422 Jmax")

BASE = "https://www.jmaxtvshop.com.ua"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "ru-RU,ru;q=0.9,uk-UA;q=0.8,en-US;q=0.7,en;q=0.6",
    "Connection": "keep-alive"
}

session = requests.Session()
session.headers.update(HEADERS)

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        if age > 3600:
            os.remove(LOCK_FILE)
            return False

        return True

    except:
        return False


def set_lock(state):

    if state:

        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:

        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)

# =========================
# LOGIN
# =========================
def login():

    print("🔐 LOGIN...", flush=True)

    login_url = BASE + "/index.php?route=account/login"

    print(
        f"🔐 LOGIN URL: {login_url}",
        flush=True
    )

    try:

        r = session.get(
            login_url,
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔐 LOGIN STATUS: {r.status_code}",
            flush=True
        )

        print(
            f"🔐 LOGIN URL FINAL: {r.url}",
            flush=True
        )

        if r.status_code != 200:

            print(
                f"❌ Страница логина: HTTP {r.status_code}",
                flush=True
            )

            return False

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        form = soup.select_one(
            'form[action*="route=account/login"]'
        )

        if not form:
            form = soup.select_one("form")

        if not form:

            print(
                "❌ Форма логина не найдена",
                flush=True
            )

            return False

        action = form.get("action")

        if not action:
            action = login_url

        if not action.startswith("http"):
            action = BASE + "/" + action.lstrip("/")

        payload = {}

        for inp in form.select("input"):

            name = inp.get("name")

            if name:
                payload[name] = inp.get("value", "")

        payload["email"] = EMAIL
        payload["password"] = PASSWORD

        print(
            f"🔐 POST LOGIN: {action}",
            flush=True
        )

        response = session.post(
            action,
            data=payload,
            headers={
                "Referer": login_url
            },
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔐 LOGIN POST STATUS: {response.status_code}",
            flush=True
        )

        print(
            f"🔐 LOGIN FINAL URL: {response.url}",
            flush=True
        )

        if response.status_code != 200:

            print(
                f"❌ LOGIN HTTP ERROR: {response.status_code}",
                flush=True
            )

            return False

        html = response.text.lower()

        if (
            "account/logout" in html
            or "route=account/logout" in html
            or "logout" in html
            or "выйти" in html
            or "вихід" in html
            or "account/account" in response.url
        ):

            print(
                "✅ LOGIN OK",
                flush=True
            )

            return True

        print(
            "⚠️ LOGIN CHECK — явного подтверждения входа нет",
            flush=True
        )

        return True

    except Exception as e:

        print(
            f"❌ LOGIN ERROR: {e}",
            flush=True
        )

        return False
        
# =========================
# STATUS
# =========================
def save_status(running=False, progress=0, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path
    }

    tmp = STATUS_PATH + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, STATUS_PATH)


# =========================
# HTTP
# =========================
def get_soup(url):

    for _ in range(3):

        try:

            r = session.get(url, timeout=30)
            
            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

        except:
            pass

        time.sleep(1)

    return BeautifulSoup("", "html.parser")

def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""


# =========================
# CATEGORIES
# =========================
def get_categories():

    print("")
    print("=" * 70)
    print("🌳 ПОИСК ВСЕХ КАТЕГОРИЙ JMAX")
    print("=" * 70)

    categories = []
    queue = []
    seen = set()

    def normalize_url(url):

        if not url:
            return ""

        url = url.strip()

        if not url:
            return ""

        if url.startswith("#"):
            return ""

        if url.startswith("javascript"):
            return ""

        if not url.startswith("http"):
            url = BASE + "/" + url.lstrip("/")

        if "route=product/category" not in url:
            return ""

        url = re.sub(
            r"[&?]page=\d+",
            "",
            url
        )

        return url

    def add_category(url):

        url = normalize_url(url)

        if not url:
            return

        if url in seen:
            return

        seen.add(url)

        categories.append(url)
        queue.append(url)

        print(
            f"📂 Категория #{len(categories)}: {url}",
            flush=True
        )

    # ==========================================================
    # ГЛАВНАЯ
    # ==========================================================

    soup = get_soup(BASE)

    if not soup or not soup.find_all(True):

        print(
            "❌ Главная страница не загрузилась",
            flush=True
        )

        return []

    print(
        "🏠 Сканируем главную...",
        flush=True
    )

    for a in soup.find_all("a", href=True):

        add_category(
            a.get("href")
        )

    # ==========================================================
    # ОБХОД КАТЕГОРИЙ
    # ==========================================================

    index = 0

    while index < len(queue):

        cat_url = queue[index]
        index += 1

        print(
            f"🌳 [{index}/{len(queue)}] {cat_url}",
            flush=True
        )

        soup = get_soup(cat_url)

        if not soup or not soup.find_all(True):
            continue

        before = len(queue)

        for a in soup.find_all("a", href=True):

            add_category(
                a.get("href")
            )

        found = len(queue) - before

        if found:

            print(
                f"   ➕ Новых категорий: {found}",
                flush=True
            )

    print("")
    print("=" * 70)
    print(
        f"🌳 ВСЕГО КАТЕГОРИЙ: {len(categories)}"
    )
    print("=" * 70)

    return categories
    
# =========================
# LAST PAGE DETECTION
# =========================
def get_last_page(soup):

    pages = [1]

    for a in soup.select("ul.pagination a"):

        href = a.get("href", "")

        m = re.search(r"[?&]page=(\d+)", href)

        if m:
            pages.append(int(m.group(1)))

    return max(pages)

# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    # первая страница без limit
    first_page = get_soup(cat_url)

    last_page = get_last_page(first_page)

    #print(f"📄 Pages: {last_page}")

    for page in range(1, last_page + 1):

        if page == 1:
            url = cat_url
        else:
            url = f"{cat_url}&page={page}"

        soup = get_soup(url)

        cards = soup.select("div.product-layout")

        #print(f"Page {page}: {len(cards)} products")

        for card in cards:

            title = ""
            sku = ""
            price = ""
            status = ""
            url_product = ""

            # TITLE + URL
            title_el = card.select_one(".us-module-title a")

            if title_el:
                title = clean(title_el.get_text())

                href = title_el.get("href", "").strip()

                if href.startswith("/"):
                    href = BASE + href

                url_product = href
           
            # SKU
            sku = ""
            
            for div in card.select("div"):
                text = clean(div.get_text())
            
                if text.startswith("Артикул"):
                    sku = text.replace("Артикул -", "").replace("Артикул:", "").strip()
                    break

            # PRICE
            price_el = card.select_one(".us-module-price-actual")

            if price_el:
                price = clean(price_el.get_text())

            # STATUS
            status = ""
            
            for span in card.select(".us-module-caption span"):
                text = clean(span.get_text())
                if text:
                    status = text
                    break

            all_items.append([
                sku,
                title,
                price,
                status,
                url_product
            ])

    return all_items
  
# =========================
# MAIN
# =========================
def run_parser():

    if is_locked():
        return

    set_lock(True)

    try:

        save_status(True, 0, USER, FILE_PATH)

        login()

        wb = Workbook()
        ws = wb.active
        ws.append(["SKU", "TITLE", "PRICE", "STATUS", "URL"])

        seen = set()

        cats = get_categories()

        if CATEGORY_LIMIT:
            cats = cats[:CATEGORY_LIMIT]

        total = len(cats)

        if total == 0:
            save_status(False, 100, USER, FILE_PATH)
            return

        for i, cat in enumerate(cats, 1):

            save_status(
                True,
                int(i / total * 100),
                USER,
                FILE_PATH
            )

            items = parse_category(cat)

            for sku, title, price, status, url in items:

                #key = sku if sku else url
                key = (title, price)

                if key in seen:
                    continue

                seen.add(key)

                if not title:
                    continue

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    url
                ])

            time.sleep(0.2)

        os.makedirs(OUTPUT_DIR, exist_ok=True)

        tmp = FILE_PATH + ".tmp"

        wb.save(tmp)

        os.replace(tmp, FILE_PATH)

        save_status(False, 100, USER, FILE_PATH)

        print("✅ Готово. Харьковская 4421-4422 Jmax")

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
