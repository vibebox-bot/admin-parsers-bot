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

print("🔥 Makita")

BASE = "https://droptools.in.ua"

# =========================
# ⚙️ SWITCH
# =========================
# CATEGORY_LIMIT = 2
CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/Makita")
FILE_PATH = os.path.join(OUTPUT_DIR, "Makita_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()
session.headers.update(HEADERS)


# =========================
# LOCK
# =========================
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

        os.makedirs(OUTPUT_DIR, exist_ok=True)

        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:

        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


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
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )

    os.replace(tmp, STATUS_PATH)


# =========================
# HTTP
# =========================
def get_soup(url):

    for _ in range(3):

        try:

            r = session.get(
                url,
                timeout=30
            )

            if r.status_code == 200:
                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

            print(
                f"⚠ HTTP {r.status_code}: {url}"
            )

        except Exception as e:

            print(
                f"⚠ Request error: {url} | {e}"
            )

        time.sleep(1)

    return BeautifulSoup(
        "",
        "html.parser"
    )


# =========================
# CLEAN
# =========================
def clean(t):

    return re.sub(
        r"\s+",
        " ",
        t
    ).strip() if t else ""


# =========================
# URL
# =========================
def absolute_url(href):

    if not href:
        return ""

    href = href.strip()

    if href.startswith("http://") or href.startswith("https://"):
        return href

    if href.startswith("/"):
        return BASE + href

    return BASE + "/" + href


# =========================
# CATEGORIES
# =========================
def get_categories():

    soup = get_soup(BASE)

    categories = []

    # Берём все ссылки меню,
    # которые ведут на ?cat=
    for a in soup.select("nav a[href*='?cat=']"):

        href = a.get("href", "").strip()

        if not href:
            continue

        href = absolute_url(href)

        if href not in categories:
            categories.append(href)

    print(
        f"📂 Categories: {len(categories)}"
    )

    for cat in categories:
        print(
            f"   {cat}"
        )

    return categories


# =========================
# LAST PAGE
# =========================
def get_last_page(soup):

    pages = [1]

    # Все ссылки пагинации
    for a in soup.select("a[href*='page=']"):

        href = a.get("href", "")

        m = re.search(
            r"[?&]page=(\d+)",
            href
        )

        if m:
            pages.append(
                int(m.group(1))
            )

    return max(pages)


# =========================
# PARSE PRODUCTS
# =========================
def parse_products(soup):

    all_items = []

    # Карточка товара
    cards = soup.select(
        "div.group.flex.flex-col.overflow-hidden.rounded-2xl.bg-white"
    )

    for card in cards:

        sku = ""
        title = ""
        price = ""
        status = ""
        url_product = ""

        # =========================
        # TITLE + URL
        # =========================

        title_el = card.select_one(
            "a[title]"
        )

        if not title_el:

            title_el = card.select_one(
                "a[href*='/product/']"
            )

        if title_el:

            title = clean(
                title_el.get_text()
            )

            href = title_el.get(
                "href",
                ""
            )

            url_product = absolute_url(
                href
            )

        # =========================
        # SKU
        # =========================

        # В карточке SKU находится
        # в span.text-xs.text-zinc-400

        sku_el = card.select_one(
            "span.text-xs.text-zinc-400"
        )

        if sku_el:

            sku = clean(
                sku_el.get_text()
            )

        # =========================
        # PRICE
        # =========================

        # Ищем текст, содержащий $
        for span in card.select(
            "span"
        ):

            text = clean(
                span.get_text()
            )

            if "$" in text:

                # Чтобы случайно не взять
                # другой текст с $
                if re.search(
                    r"\d",
                    text
                ):

                    price = text
                    break

        # =========================
        # STATUS
        # =========================

        # В наличии
        for span in card.select(
            "span"
        ):

            text = clean(
                span.get_text()
            )

            if (
                "В наявності" in text
                or "Немає в наявності" in text
                or "Немає" in text
            ):

                status = text
                break

        # =========================
        # SAVE
        # =========================

        if title:

            all_items.append([
                sku,
                title,
                price,
                status,
                url_product
            ])

    return all_items


# =========================
# PARSE MAIN PAGE
# =========================
def parse_main():

    print(
        "🏠 Парсим главную..."
    )

    soup = get_soup(BASE)

    items = parse_products(
        soup
    )

    print(
        f"   Товаров на главной: {len(items)}"
    )

    return items


# =========================
# PARSE CATEGORY
# =========================
def parse_category(cat_url):

    all_items = []

    # Первая страница
    first_page = get_soup(
        cat_url
    )

    last_page = get_last_page(
        first_page
    )

    print(
        f"📂 {cat_url}"
    )

    print(
        f"   📄 Страниц: {last_page}"
    )

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:

            url = cat_url

        else:

            separator = (
                "&"
                if "?" in cat_url
                else "?"
            )

            url = (
                f"{cat_url}"
                f"{separator}"
                f"page={page}"
            )

        soup = get_soup(
            url
        )

        items = parse_products(
            soup
        )

        all_items.extend(
            items
        )

        print(
            f"      page {page}: "
            f"{len(items)} товаров"
        )

    return all_items


# =========================
# MAIN
# =========================
def run_parser():

    if is_locked():
        print(
            "⚠ Парсер уже запущен"
        )
        return

    set_lock(True)

    try:

        save_status(
            True,
            0,
            USER,
            FILE_PATH
        )

        # =========================
        # EXCEL
        # =========================

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

        seen = set()

        # =========================
        # MAIN PAGE
        # =========================

        main_items = parse_main()

        for item in main_items:

            sku, title, price, status, url = item

            if sku:
                key = ("sku", sku)
            elif url:
                key = ("url", url)
            else:
                key = (
                    "title",
                    title
                )

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

        # =========================
        # CATEGORIES
        # =========================

        cats = get_categories()

        if CATEGORY_LIMIT:
            cats = cats[
                :CATEGORY_LIMIT
            ]

        total = len(cats)

        if total == 0:

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return

        # =========================
        # PARSE CATEGORIES
        # =========================

        for i, cat in enumerate(
            cats,
            1
        ):

            # Главная занимает 0–10%
            progress = 10 + int(
                i / total * 90
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )

            items = parse_category(
                cat
            )

            for item in items:

                sku, title, price, status, url = item

                # =========================
                # DEDUPLICATION
                # =========================

                if sku:

                    key = (
                        "sku",
                        sku
                    )

                elif url:

                    key = (
                        "url",
                        url
                    )

                else:

                    key = (
                        "title",
                        title
                    )

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

        # =========================
        # SAVE
        # =========================

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        tmp = FILE_PATH + ".tmp"

        wb.save(
            tmp
        )

        os.replace(
            tmp,
            FILE_PATH
        )

        print(
            f"📦 Всего уникальных товаров: "
            f"{len(seen)}"
        )

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        print(
            "✅ Готово. Makita"
        )

    finally:

        set_lock(False)


# =========================
# START
# =========================
if __name__ == "__main__":

    run_parser()
