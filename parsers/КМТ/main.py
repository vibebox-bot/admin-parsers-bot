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

print("🔥 Харьковская КМТ")


BASE = "https://kmt5.com.ua"


# =========================
# ⚙️ SWITCH
# =========================

CATEGORY_LIMIT = 1
# CATEGORY_LIMIT = None


EMAIL = "YOUR_EMAIL"
PASSWORD = "YOUR_PASSWORD"


OUTPUT_DIR = os.path.abspath("output/КМТ")
FILE_PATH = os.path.join(OUTPUT_DIR, "КМТ_LIVE.xlsx")
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
# LOGIN
# =========================

def login():

    print("🔐 Авторизация...")

    session.get(BASE)

    login_url = BASE + "/login/?ajax=1"

    payload = {
        "email": EMAIL,
        "password": PASSWORD
    }

    try:

        r = session.post(
            login_url,
            data=payload,
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "Referer": BASE + "/access-denied"
            },
            timeout=30
        )

        print("LOGIN:", r.status_code)

        try:
            print(r.json())
        except:
            print(r.text[:500])

    except Exception as e:

        print("❌ Ошибка LOGIN:", e)
        return False


    check = session.get(BASE)

    text = check.text.lower()

    if (
        "logout" in text
        or "выйти" in text
        or "личный кабинет" in text
    ):

        print("✅ LOGIN OK")
        return True

    print("⚠ LOGIN CHECK")
    return False


# =========================
# STATUS
# =========================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path=""
):

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        "file_path": file_path
    }

    tmp = STATUS_PATH + ".tmp"

    with open(
        tmp,
        "w",
        encoding="utf-8"
    ) as f:

        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2
        )

    os.replace(
        tmp,
        STATUS_PATH
    )


# =========================
# HTTP
# =========================

def get_soup(url):

    for attempt in range(3):

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
                f"⚠ Ошибка запроса "
                f"{attempt + 1}/3: {e}"
            )

        time.sleep(1)

    return BeautifulSoup(
        "",
        "html.parser"
    )


def clean(t):

    return (
        re.sub(r"\s+", " ", t).strip()
        if t else ""
    )


def absolute_url(url):

    if not url:
        return ""

    url = url.strip()

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("/"):
        return BASE + url

    if url.startswith("http://"):
        return url

    if url.startswith("https://"):
        return url

    return BASE + "/" + url.lstrip("/")


# =========================
# CATEGORIES
# =========================

def get_categories():

    print("")
    print("=" * 70)
    print("📂 СОБИРАЕМ ДЕРЕВО КАТЕГОРИЙ")
    print("=" * 70)

    soup = get_soup(BASE)

    cats = []
    seen = set()

    menu = soup.select_one(
        "nav.menu-left > ul"
    )

    if not menu:

        print(
            "❌ nav.menu-left > ul "
            "не найдено"
        )

        return cats


    def walk_menu(ul, level=0):

        for li in ul.find_all(
            "li",
            recursive=False
        ):

            a = li.find(
                "a",
                href=True,
                recursive=False
            )

            if a:

                href = absolute_url(
                    a.get("href")
                )

                name = clean(
                    a.get_text(
                        " ",
                        strip=True
                    )
                )

                if (
                    href
                    and href.startswith(BASE)
                    and href not in seen
                ):

                    seen.add(href)

                    cats.append({
                        "name": name,
                        "url": href,
                        "level": level
                    })

                    print(
                        "   "
                        + "  " * level
                        + f"📂 {name}"
                        + f" → {href}"
                    )


            child_ul = li.find(
                "ul",
                recursive=False
            )

            if child_ul:

                walk_menu(
                    child_ul,
                    level + 1
                )


    walk_menu(menu)

    print("")
    print(
        f"📂 ВСЕГО КАТЕГОРИЙ: {len(cats)}"
    )

    return cats


# =========================
# FIND SUBCATEGORIES
# =========================

def get_subcategories(
    soup,
    current_url,
    known_products
):

    result = []
    seen = set()


    # --------------------------------
    # Варианты блоков категорий
    # --------------------------------

    selectors = [

        ".catalog-group > li a",
        ".catalog-group_item a",

        ".catalog-category a",

        ".category-list a",

        ".sub-category a",

        ".categories-list a",

        ".box-category a",

    ]


    for selector in selectors:

        for a in soup.select(selector):

            href = absolute_url(
                a.get("href")
            )

            if not href:
                continue

            if not href.startswith(BASE):
                continue

            if href == current_url:
                continue

            if href in known_products:
                continue

            if href in seen:
                continue


            # --------------------------------
            # Не принимаем ссылки на товары
            # --------------------------------

            parent_product = a.closest(
                ".list-catalog_item"
            )

            if parent_product:
                continue


            # --------------------------------
            # Не берём очевидные служебные URL
            # --------------------------------

            bad_parts = [
                "javascript:",
                "#",
                "/login",
                "/checkout",
                "/account",
                "/cart"
            ]

            if any(
                x in href.lower()
                for x in bad_parts
            ):
                continue


            seen.add(href)
            result.append(href)


    return result


# =========================
# PARSE PRODUCT
# =========================

def parse_product(
    url,
    status
):

    soup = get_soup(url)

    # =========================
    # TITLE
    # =========================

    title = ""

    h1 = soup.select_one("h1")

    if h1:

        title = clean(
            h1.get_text()
        )


    # =========================
    # SKU
    # =========================

    sku = ""


    # =========================
    # CODE
    # =========================

    code = ""

    box = soup.select_one(
        ".box-card_code"
    )

    if box:

        text = box.get_text(
            " ",
            strip=True
        )


        m = re.search(
            r"Код товара:\s*(Ц-\d+)",
            text
        )

        if m:

            sku = m.group(1)


        m = re.search(
            r"Код:\s*(\S+)",
            text
        )

        if m:

            code = m.group(1)


    # =========================
    # PRICE
    # =========================

    price = ""

    new_price = soup.select_one(
        ".price__new"
    )

    if new_price:

        price = clean(
            new_price.get_text()
        )

    else:

        box_price = soup.select_one(
            ".box-card_hryvnia"
        )

        if box_price:

            price = clean(
                box_price.get_text()
            )


    return [
        sku,
        code,
        title,
        price,
        status,
        url
    ]


# =========================
# PARSE CATEGORY
# =========================

def parse_category(
    cat_url,
    visited_categories=None,
    global_products=None
):

    if visited_categories is None:

        visited_categories = set()


    if global_products is None:

        global_products = set()


    # =========================
    # CATEGORY DUPLICATE
    # =========================

    if cat_url in visited_categories:

        return []


    visited_categories.add(
        cat_url
    )


    result = []

    seen_pages = set()


    print("")
    print("=" * 70)
    print("📂 CATEGORY")
    print(cat_url)
    print("=" * 70)


    # ==========================================================
    # 1. СТРАНИЦЫ КАТЕГОРИИ
    # ==========================================================

    page = 1

    while True:

        # --------------------------------
        # URL страницы
        # --------------------------------

        if page == 1:

            url = cat_url

        else:

            sep = (
                "&"
                if "?" in cat_url
                else "?"
            )

            url = (
                f"{cat_url}"
                f"{sep}page={page}&ajax=1"
            )


        if url in seen_pages:

            print(
                "⛔ Страница уже была"
            )

            break


        seen_pages.add(url)


        print("")
        print(
            f"📄 PAGE {page}"
        )

        print(url)


        soup = get_soup(url)


        # --------------------------------
        # Карточки
        # --------------------------------

        cards = soup.select(
            "div.list-catalog_item"
        )


        print(
            f"   Карточек на странице: "
            f"{len(cards)}"
        )


        if not cards:

            print(
                "   ⛔ Товаров больше нет"
            )

            break


        added = 0


        for card in cards:

            a = card.select_one(
                ".list-catalog_title a"
            )

            if not a:
                continue


            href = absolute_url(
                a.get("href")
            )


            if not href:
                continue


            # --------------------------------
            # URL = уникальный товар
            # --------------------------------

            if href in global_products:

                continue


            global_products.add(
                href
            )


            # --------------------------------
            # STATUS
            # --------------------------------

            status = ""

            label = card.select_one(
                ".product__label"
            )

            if label:

                status = clean(
                    label.get_text()
                )


            # --------------------------------
            # PRODUCT
            # --------------------------------

            item = parse_product(
                href,
                status
            )


            if not item[2]:

                continue


            result.append(item)

            added += 1


        print(
            f"   ➕ Добавлено: {added}"
        )


        # --------------------------------
        # Кнопка "Показать ещё"
        # --------------------------------

        next_button = soup.select_one(
            ".btn__more a"
        )


        if not next_button:

            print(
                "   ⛔ Кнопки следующей "
                "страницы нет"
            )

            break


        next_url = absolute_url(
            next_button.get("href")
        )


        if not next_url:

            break


        if next_url in seen_pages:

            print(
                "   ⛔ Следующая страница "
                "уже была"
            )

            break


        page += 1

        time.sleep(0.3)


    # ==========================================================
    # 2. ИЩЕМ ПОДКАТЕГОРИИ
    # ==========================================================

    print("")
    print(
        "🔎 Проверяем подкатегории..."
    )


    # Берём первую страницу категории
    # ещё раз, чтобы искать дерево

    category_soup = get_soup(
        cat_url
    )


    subcategories = get_subcategories(
        category_soup,
        cat_url,
        global_products
    )


    print(
        f"   📂 Найдено возможных "
        f"подкатегорий: "
        f"{len(subcategories)}"
    )


    # ==========================================================
    # 3. РЕКУРСИВНО ИДЁМ ВНИЗ
    # ==========================================================

    for sub_url in subcategories:

        if sub_url in visited_categories:

            continue


        print("")
        print(
            "➡️ ПЕРЕХОД В ПОДКАТЕГОРИЮ:"
        )

        print(sub_url)


        sub_items = parse_category(
            sub_url,
            visited_categories,
            global_products
        )


        result.extend(
            sub_items
        )


        time.sleep(0.3)


    return result


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
        # LOGIN
        # =========================

        login()


        # =========================
        # ВАЖНО:
        #
        # set_product_group()
        #
        # НЕ ВЫЗЫВАЕМ
        #
        # product_group=1
        # больше НЕ включаем
        # =========================


        # =========================
        # EXCEL
        # =========================

        wb = Workbook()

        ws = wb.active

        ws.title = "Товары"


        ws.append([
            "SKU",
            "CODE",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])


        # =========================
        # CATEGORIES
        # =========================

        cats = get_categories()


        print("")
        print(
            f"📂 Категорий найдено: "
            f"{len(cats)}"
        )


        if CATEGORY_LIMIT:

            cats = cats[
                :CATEGORY_LIMIT
            ]

            print(
                f"🧪 ТЕСТОВЫЙ РЕЖИМ: "
                f"{len(cats)} категория"
            )


        total = len(cats)


        if total == 0:

            print(
                "❌ Категории не найдены"
            )

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return


        # =========================
        # GLOBAL PRODUCT SET
        # =========================

        global_products = set()

        visited_categories = set()


        # =========================
        # PARSE
        # =========================

        total_items = 0


        for i, cat in enumerate(
            cats,
            1
        ):


            progress = int(
                (i / total) * 100
            )


            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )


            print("")
            print("")
            print(
                "#" * 70
            )

            print(
                f"📂 КАТЕГОРИЯ "
                f"{i}/{total}"
            )

            print(
                f"📂 {cat['name']}"
            )

            print(
                f"🔗 {cat['url']}"
            )

            print(
                "#" * 70
            )


            items = parse_category(
                cat["url"],
                visited_categories,
                global_products
            )


            print("")
            print(
                f"📦 Товаров найдено "
                f"в этой ветке: "
                f"{len(items)}"
            )


            # =========================
            # WRITE EXCEL
            # =========================

            for (
                sku,
                code,
                title,
                price,
                status,
                url
            ) in items:


                if not title:

                    continue


                ws.append([
                    sku,
                    code,
                    title,
                    price,
                    status,
                    url
                ])


                total_items += 1


            print(
                f"📦 ВСЕГО УНИКАЛЬНЫХ "
                f"ТОВАРОВ: "
                f"{len(global_products)}"
            )


            time.sleep(0.2)


        # =========================
        # SAVE
        # =========================

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )


        tmp = FILE_PATH + ".tmp"


        wb.save(tmp)


        os.replace(
            tmp,
            FILE_PATH
        )


        # =========================
        # DONE
        # =========================

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )


        print("")
        print("=" * 70)

        print(
            "✅ ГОТОВО"
        )

        print(
            f"📦 Всего товаров: "
            f"{total_items}"
        )

        print(
            f"📂 Категорий обработано: "
            f"{len(visited_categories)}"
        )

        print(
            f"📄 Файл: "
            f"{FILE_PATH}"
        )

        print("=" * 70)


    finally:

        set_lock(False)


# =========================
# START
# =========================

if __name__ == "__main__":

    run_parser()
