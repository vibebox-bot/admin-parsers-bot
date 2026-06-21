import os
import json
import time
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

# =========================
# OUTPUT CONFIG (ТВОЙ ТЕСТ)
# =========================
OUTPUT_DIR = os.path.abspath("output/hi_tech_test")

FILE_PATH = os.path.join(OUTPUT_DIR, "hi_tech_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


BASE_URL = "https://hi-tech-odessa.com.ua"

TEST_CATEGORY = "https://hi-tech-odessa.com.ua/?product_cat=bluetooth-%d0%ba%d0%be%d0%bb%d0%be%d0%bd%d0%ba%d0%b8-%d0%bf%d0%be%d1%80%d1%82%d0%b0%d1%82%d0%b8%d0%b2%d0%bd%d1%8b%d0%b5"


# =========================
# STATUS
# =========================
def save_status(progress=0, running=True):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "time": time.strftime("%Y-%m-%d %H:%M:%S")
    }

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


# =========================
# HTTP
# =========================
def get_soup(url):
    r = requests.get(url, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


# =========================
# CATEGORY PAGES
# =========================
def get_products_links(category_url):
    links = []

    page = 1

    while True:
        url = category_url + f"&paged={page}" if "paged=" not in category_url else category_url

        soup = get_soup(url)

        items = soup.select("li.product a[href*='product=']")

        if not items:
            break

        for i in items:
            href = i.get("href")
            if href:
                links.append(href)

        # pagination check
        next_btn = soup.select_one("a.next.page-numbers")
        if not next_btn:
            break

        page += 1
        save_status(progress=min(page * 5, 90))

    return list(set(links))


# =========================
# PRODUCT PARSER
# =========================
def parse_product(url):
    soup = get_soup(url)

    # name
    title = soup.select_one(".product_title.entry-title")
    title = title.get_text(strip=True) if title else ""

    # sku
    sku = soup.select_one(".sku")
    sku = sku.get_text(strip=True) if sku else ""

    # price
    price_tag = soup.select_one(".woocommerce-Price-amount")
    price = price_tag.get_text(strip=True) if price_tag else ""

    # stock
    stock = "in_stock"
    if soup.select_one("p.out-of-stock"):
        stock = "out_of_stock"

    return {
        "url": url,
        "title": title,
        "sku": sku,
        "price": price,
        "stock": stock
    }


# =========================
# MAIN
# =========================
def main():
    print("🚀 HI-TECH PARSER STARTED")

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    save_status(0, True)

    # 1) берем товары из категории
    product_links = get_products_links(TEST_CATEGORY)

    total = len(product_links)
    print(f"📦 FOUND PRODUCTS: {total}")

    data = []

    # 2) парсим карточки
    for idx, url in enumerate(product_links, start=1):
        try:
            item = parse_product(url)
            data.append(item)

        except Exception as e:
            print("ERROR:", url, e)

        progress = int((idx / total) * 100)
        save_status(progress=progress, running=True)

        print(f"[{progress}%] {url}")

        time.sleep(0.5)

    # 3) финал
    save_status(progress=100, running=False)

    print("✅ DONE")
    print("TOTAL:", len(data))

    # пока просто json (xlsx подключим дальше)
    with open(os.path.join(OUTPUT_DIR, "data.json"), "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
