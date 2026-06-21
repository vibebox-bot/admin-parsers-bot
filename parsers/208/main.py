import os
import json
import time
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook

# =========================
# OUTPUT PATHS (ТЕСТ)
# =========================

OUTPUT_DIR = os.path.abspath("output/208")

FILE_PATH = os.path.join(OUTPUT_DIR, "Харьковская_208_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

os.makedirs(OUTPUT_DIR, exist_ok=True)


BASE_URL = "https://hi-tech-odessa.com.ua"

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


# =========================
# STATUS
# =========================

def save_status(data):
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_status():
    if not os.path.exists(STATUS_PATH):
        return None
    try:
        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return None


# =========================
# PARSING HELPERS
# =========================

def get_soup(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


def parse_product(url):
    soup = get_soup(url)

    # название
    title = soup.select_one("h1.product_title")
    title = title.text.strip() if title else ""

    # SKU
    sku = soup.select_one(".sku")
    sku = sku.text.strip() if sku else ""

    # цена
    price = soup.select_one(".woocommerce-Price-amount")
    price = price.text.strip() if price else ""

    # наличие
    stock = soup.select_one("p.stock")
    if stock:
        stock_text = stock.text.strip()
    else:
        stock_text = "В наличии"

    return {
        "title": title,
        "sku": sku,
        "price": price,
        "stock": stock_text,
        "url": url
    }


def get_products_from_category(category_url, max_pages=1):
    products = []

    for page in range(1, max_pages + 1):
        url = f"{category_url}&paged={page}" if "?" in category_url else f"{category_url}?paged={page}"

        soup = get_soup(url)

        links = soup.select("a.woocommerce-LoopProduct-link, a.woocommerce-loop-product__link")

        if not links:
            break

        for a in links:
            href = a.get("href")
            if href and href not in [p["url"] for p in products]:
                products.append({"url": href})

    return products


# =========================
# MAIN PARSER
# =========================

def run():
    category_url = "https://hi-tech-odessa.com.ua/?product_cat=bluetooth-%d0%b0%d0%ba%d1%83%d1%81%d1%82%d0%b8%d0%ba%d0%b0-%d0%b0%d0%ba%d0%ba%d1%83%d0%bc%d1%83%d0%bb%d1%8f%d1%82%d0%be%d1%80%d0%bd%d0%b0%d1%8f"

    print("🚀 START PARSER HI-TECH TEST")

    # ⚠️ защита от дубля
    if os.path.exists(LOCK_FILE):
        print("⛔ ALREADY RUNNING")
        return

    open(LOCK_FILE, "w").close()

    wb = Workbook()
    ws = wb.active
    ws.append(["Название", "SKU", "Цена", "Наличие", "URL"])

    save_status({
        "running": True,
        "progress": 0,
        "done": 0,
        "total": 0
    })

    product_links = get_products_from_category(category_url, max_pages=2)

    total = len(product_links)

    save_status({
        "running": True,
        "progress": 0,
        "done": 0,
        "total": total
    })

    done = 0

    for item in product_links:

        if os.path.exists(LOCK_FILE) is False:
            break

        try:
            data = parse_product(item["url"])

            ws.append([
                data["title"],
                data["sku"],
                data["price"],
                data["stock"],
                data["url"]
            ])

            done += 1
            progress = int(done / total * 100)

            save_status({
                "running": True,
                "progress": progress,
                "done": done,
                "total": total
            })

            print(f"[{progress}%] {data['title']}")

            # 💾 ВАЖНО — сохраняем ПОСТОЯННО
            wb.save(FILE_PATH)

            time.sleep(0.3)

        except Exception as e:
            print("ERROR:", e)

    wb.save(FILE_PATH)

    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)

    save_status({
        "running": False,
        "progress": 100,
        "done": done,
        "total": total
    })

    print("✅ DONE")


# =========================
# ENTRY FOR run.py
# =========================

def main():
    run()
