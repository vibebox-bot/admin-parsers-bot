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

print("🔥 DELLTA LIFE PARSER")

BASE = "https://b2b.delltalife.com"

# =========================
# ⚙️ SWITCH
# =========================
CATEGORY_LIMIT = 1
#CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"

OUTPUT_DIR = os.path.abspath("output/Dellta")
FILE_PATH = os.path.join(OUTPUT_DIR, "Dellta_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()
session.headers.update(HEADERS)

# =========================
# LOGIN
# =========================
def login():
    print("LOGIN...")

    login_url = BASE + "/login"

    # 1. GET страницу (ВАЖНО для cookies + возможных токенов)
    r = session.get(login_url)
    soup = BeautifulSoup(r.text, "html.parser")

    # 2. собираем hidden поля (если есть csrf / token)
    payload = {}

    for inp in soup.select("form input"):
        name = inp.get("name")
        if name:
            payload[name] = inp.get("value", "")

    # 3. подставляем логин/пароль (ВАЖНО: правильные name из HTML)
    payload["email_auth"] = EMAIL
    payload["pass_auth"] = PASSWORD

    # иногда нужно:
    payload["remember"] = "on"

    # 4. отправляем именно AJAX endpoint (как в форме)
    login_action = BASE + "/themes/default/ajax/login.php"

    r2 = session.post(
        login_action,
        data=payload,
        headers={
            "X-Requested-With": "XMLHttpRequest",
            "Referer": login_url
        }
    )

    print("STATUS:", r2.status_code)
    print("RESPONSE:", r2.text[:300])

    # 5. проверка успеха (очень важно)
    if "error" not in r2.text.lower():
        print("LOGIN OK")
    else:
        print("LOGIN FAILED")

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
    r = session.get(url, timeout=30)
    return BeautifulSoup(r.text, "html.parser")


def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""


# =========================
# CATEGORIES
# =========================
def get_categories():
    soup = get_soup(BASE)

    cats = []

    for a in soup.select("a[href*='catalog'], a[href*='category'], .catalog a"):
        href = a.get("href")

        if not href:
            continue

        if href.startswith("/"):
            href = BASE + href

        if href not in cats:
            cats.append(href)

    return cats


# =========================
# LAST PAGE DETECTION
# =========================
def get_last_page(soup):
    pages = []

    for a in soup.select(".pagination .page-link[pn]"):
        pn = a.get("pn")
        if pn and pn.isdigit():
            pages.append(int(pn))

    return max(pages) if pages else 1


# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):
    print("HTML SIZE:", len(soup.text))
    
    open("cat_debug.html", "w", encoding="utf-8").write(str(soup))
    
    print("CATEGORY:", cat_url)

    all_items = []

    first_page = get_soup(cat_url)

    last_page = get_last_page(first_page)

    print("PAGES:", last_page)

    for page in range(1, last_page + 1):

        if page == 1:
            soup = first_page
        else:
            url = f"{cat_url}?page={page}"
            soup = get_soup(url)

        cards = soup.select(".ItemDiv")

        for card in cards:

            title = clean(card.select_one(".item-name") and card.select_one(".item-name").text)

            sku = clean(card.select_one(".sku-block .gray") and card.select_one(".sku-block .gray").text)

            status = clean(card.select_one(".sku-block .are-available") and card.select_one(".sku-block .are-available").text)

            price = ""

            for row in card.select(".price-table tr"):
                tds = row.select("td")
                if len(tds) >= 2:
                    name = clean(tds[0].text)
                    if "Комп. ДИЛЕР" in name:
                        price = clean(tds[1].text)

            url_tag = card.select_one(".item-name")
            url = ""
            if url_tag and url_tag.get("href"):
                url = url_tag.get("href")
                if url.startswith("/"):
                    url = BASE + url

            all_items.append([sku, title, price, status, url])

    return all_items


# =========================
# MAIN
# =========================
def run_parser():

    save_status(True, 0, USER, FILE_PATH)

    login()

    wb = Workbook()
    ws = wb.active
    ws.append(["SKU", "TITLE", "PRICE", "STATUS", "URL"])

    seen = set()

    cats = get_categories()

    print("CATEGORIES:", len(cats))

    if CATEGORY_LIMIT:
        cats = cats[:CATEGORY_LIMIT]

    total = len(cats)

    for i, cat in enumerate(cats, 1):

        save_status(True, int(i / total * 100), USER, FILE_PATH)

        items = parse_category(cat)

        for sku, title, price, status, url in items:

            key = sku if sku else url

            if key in seen:
                continue

            seen.add(key)

            if not title:
                continue

            ws.append([sku, title, price, status, url])

        print(f"DONE CATEGORY {i}/{total} -> {len(items)} items")

        time.sleep(0.3)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    tmp = FILE_PATH + ".tmp"
    wb.save(tmp)
    os.replace(tmp, FILE_PATH)

    save_status(False, 100, USER, FILE_PATH)

    print("DONE")


if __name__ == "__main__":
    run_parser()
