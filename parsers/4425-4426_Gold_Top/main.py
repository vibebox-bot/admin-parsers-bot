import os
import json
import time
import requests
import pandas as pd
from bs4 import BeautifulSoup
from urllib.parse import urljoin


# =========================
# CONFIG
# =========================
BASE_URL = "https://www.gold-tor.com.ua"

CATEGORY_LIMIT = 1  # None = все категории

BASE_DIR = os.path.abspath("output/4425-4426_Gold_Top")

FILE_PATH = os.path.join(BASE_DIR, "Харьковская_4425-4426_Gold_Top_LIVE.xlsx")
STATUS_PATH = os.path.join(BASE_DIR, "status.json")
LOCK_FILE = os.path.join(BASE_DIR, "lock.txt")


# =========================
# SESSION (важно для логина)
# =========================
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0"
})


# =========================
# STATUS
# =========================
def save_status(text):
    os.makedirs(BASE_DIR, exist_ok=True)
    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump({"status": text}, f, ensure_ascii=False, indent=2)


# =========================
# LOCK
# =========================
def lock():
    os.makedirs(BASE_DIR, exist_ok=True)
    with open(LOCK_FILE, "w") as f:
        f.write("running")


def unlock():
    if os.path.exists(LOCK_FILE):
        os.remove(LOCK_FILE)


# =========================
# LOGIN (OpenCart-safe)
# =========================
def login():
    login_url = BASE_URL + "/index.php?route=account/login"

    r = session.get(login_url)
    soup = BeautifulSoup(r.text, "html.parser")

    payload = {}

    # забираем hidden поля (csrf/token если есть)
    for inp in soup.select("form input"):
        name = inp.get("name")
        value = inp.get("value", "")
        if name:
            payload[name] = value

    # твои данные
    payload.update({
        "email": "Sawrun_05@icloud.com",
        "password": "18022021"
    })

    r2 = session.post(login_url, data=payload)

    if "logout" in r2.text.lower() or "account/logout" in r2.text.lower():
        print("✅ LOGIN OK")
        return True

    print("❌ LOGIN FAILED")
    return False


# =========================
# CATEGORIES
# =========================
def get_categories():
    r = session.get(BASE_URL)
    soup = BeautifulSoup(r.text, "html.parser")

    cats = []

    for a in soup.select("#d_category_menu_list a.link-level-1"):
        cats.append({
            "title": a.get_text(strip=True),
            "url": urljoin(BASE_URL, a["href"])
        })

    return cats


# =========================
# PRODUCT CARD
# =========================
def parse_card(card):
    a = card.select_one(".product-name a")
    if not a:
        return None

    title = a.get_text(strip=True)
    url = urljoin(BASE_URL, a["href"])

    img = card.select_one("img")
    image = img["src"] if img else None

    price_el = card.select_one(".price")
    price = price_el.get_text(" ", strip=True) if price_el else None

    return {
        "title": title,
        "url": url,
        "image": image,
        "price": price
    }


# =========================
# PAGE PARSER
# =========================
def parse_page(url):
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    items = []

    for card in soup.select(".product-item"):
        p = parse_card(card)
        if p:
            items.append(p)

    return items


# =========================
# PAGINATION
# =========================
def get_pages(url):
    r = session.get(url)
    soup = BeautifulSoup(r.text, "html.parser")

    max_page = 1

    for a in soup.select(".pagination a"):
        href = a.get("href", "")
        if "page=" in href:
            try:
                p = int(href.split("page=")[-1])
                max_page = max(max_page, p)
            except:
                pass

    return max_page


# =========================
# CATEGORY PARSER
# =========================
def parse_category(cat):
    print(f"\n📦 {cat['title']}")

    max_page = get_pages(cat["url"])

    all_products = []

    for page in range(1, max_page + 1):
        url = cat["url"] if page == 1 else f"{cat['url']}?page={page}"

        print("  page:", page)

        items = parse_page(url)
        all_products.extend(items)

        time.sleep(0.3)

    return all_products


# =========================
# SAVE EXCEL
# =========================
def save_excel(data):
    os.makedirs(BASE_DIR, exist_ok=True)

    df = pd.DataFrame(data)
    df.to_excel(FILE_PATH, index=False)

    print(f"\n💾 EXCEL SAVED: {FILE_PATH}")


# =========================
# MAIN
# =========================
def main():
    lock()
    save_status("start")

    # 1. login
    if not login():
        save_status("login_failed")
        unlock()
        return

    # 2. categories
    cats = get_categories()

    if CATEGORY_LIMIT is not None:
        cats = cats[:CATEGORY_LIMIT]

    print(f"📂 categories: {len(cats)}")

    all_data = []

    # 3. parse
    for cat in cats:
        products = parse_category(cat)
        all_data.extend(products)

    # 4. save
    save_excel(all_data)

    save_status("done")
    unlock()


if __name__ == "__main__":
    main()
