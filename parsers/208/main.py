import os
import time
import json
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# =========================
# OUTPUT (TEST MODE)
# =========================
OUTPUT_DIR = os.path.abspath("output/hi_tech_test")

FILE_PATH = os.path.join(OUTPUT_DIR, "hi_tech_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)


CATEGORY_URL = "https://hi-tech-odessa.com.ua/?product_cat=bluetooth-%d0%b0%d0%ba%d1%83%d1%81%d1%82%d0%b8%d0%ba%d0%b0-%d0%b0%d0%ba%d0%ba%d1%83%d0%bc%d1%83%d0%bb%d1%8f%d1%82%d0%be%d1%80%d0%bd%d0%b0%d1%8f"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

# =========================
# STATUS
# =========================
def save_status(data):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_page(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return BeautifulSoup(r.text, "html.parser")


# =========================
# LOCK
# =========================
def create_lock():
    if os.path.exists(LOCK_FILE):
        age = time.time() - os.path.getmtime(LOCK_FILE)
        if age > 300:
            os.remove(LOCK_FILE)
        else:
            return False

    open(LOCK_FILE, "w").close()
    return True


def remove_lock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# GET PRODUCT LINKS (ВАЖНО)
# =========================
def get_products():
    links = []
    page = 1

    while True:
        url = CATEGORY_URL + f"&paged={page}"

        try:
            r = requests.get(url, headers=HEADERS, timeout=20)

            # 🔥 СТОП НА 404
            if r.status_code == 404:
                break

            soup = BeautifulSoup(r.text, "html.parser")

            items = soup.select("ul.products li.product a")

            # 🔥 если товаров нет — стоп
            if not items:
                break

            for i in items:
                href = i.get("href")

                # берем только товары
                if href and "?product=" in href:
                    links.append(href)

            print(f"📄 PAGE {page} OK - {len(items)} items")

        except Exception as e:
            print("PAGE ERROR:", e)
            break

        page += 1

        if page > 100:
            break

    return list(set(links))


# =========================
# PARSE PRODUCT
# =========================
def parse_product(url):
    soup = load_page(url)

    title = soup.select_one(".product_title.entry-title")
    title = title.text.strip() if title else "-"

    sku = soup.select_one(".sku")
    sku = sku.text.strip() if sku else "-"

    price = soup.select_one(".woocommerce-Price-amount")
    price = price.text.strip() if price else "-"

    stock = "В наличии"
    if soup.select_one(".out-of-stock"):
        stock = "Нет в наличии"

    return {
        "title": title,
        "sku": sku,
        "price": price,
        "stock": stock,
        "url": url
    }


# =========================
# MAIN RUN
# =========================
def run_parser():

    print("🚀 START PARSER HI-TECH TEST")

    if not create_lock():
        print("⛔ ALREADY RUNNING")
        return

    try:
        wb = Workbook()
        ws = wb.active
        ws.append(["Название", "SKU", "Цена", "Наличие", "URL"])

        links = get_products()
        total = len(links)

        print(f"📦 FOUND PRODUCTS: {total}")

        if total == 0:
            save_status({"running": False, "progress": 0, "error": "NO PRODUCTS FOUND"})
            return

        done = 0

        save_status({
            "running": True,
            "progress": 0,
            "done": 0,
            "total": total
        })

        for url in links:

            data = parse_product(url)

            ws.append([
                data["title"],
                data["sku"],
                data["price"],
                data["stock"],
                data["url"]
            ])

            done += 1
            progress = int(done / total * 100)

            print(f"[{progress}%] {data['title']}")

            save_status({
                "running": True,
                "progress": progress,
                "done": done,
                "total": total
            })

            wb.save(FILE_PATH)
            time.sleep(0.3)

        wb.save(FILE_PATH)

        save_status({
            "running": False,
            "progress": 100,
            "done": done,
            "total": total
        })

        print("✅ DONE")

    except Exception as e:
        print("❌ ERROR:", e)

        save_status({
            "running": False,
            "progress": 0,
            "error": str(e)
        })

    finally:
        remove_lock()
