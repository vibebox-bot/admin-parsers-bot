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

# ==========================================================
# ⚙️ НАСТРОЙКИ
# ==========================================================

BASE = "https://www.jmaxtvshop.com.ua"

# ==========================================================
# ⚙️ SWITCH
# ==========================================================

# CATEGORY_LIMIT = 1
CATEGORY_LIMIT = None

# ==========================================================
# 📁 FILES
# ==========================================================

OUTPUT_DIR = os.path.abspath(
    "output/4421-4422_Jmax"
)

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "4421-4422_Jmax_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "lock.txt"
)

# ==========================================================
# 🌐 HEADERS
# ==========================================================

HEADERS = {
    "User-Agent":
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36",

    "Accept":
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,image/avif,"
        "image/webp,*/*;q=0.8",

    "Accept-Language":
        "ru-RU,ru;q=0.9,uk;q=0.8,en-US;q=0.7,en;q=0.6",

    "Connection":
        "keep-alive"
}

# ==========================================================
# SESSION
# ==========================================================

session = requests.Session()

session.headers.update(
    HEADERS
)


# ==========================================================
# LOCK
# ==========================================================

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:

        age = (
            time.time()
            -
            os.path.getmtime(LOCK_FILE)
        )

        if age > 3600:

            os.remove(LOCK_FILE)

            return False

        return True

    except:

        return False


def set_lock(state):

    if state:

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        with open(
            LOCK_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            f.write(
                str(time.time())
            )

    else:

        if os.path.exists(
            LOCK_FILE
        ):

            os.remove(
                LOCK_FILE
            )


# ==========================================================
# STATUS
# ==========================================================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path=""
):

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    data = {

        "running":
            running,

        "progress":
            progress,

        "user":
            user,

        "time":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M:%S"
            ),

        "file_path":
            file_path
    }

    tmp = (
        STATUS_PATH +
        ".tmp"
    )

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


# ==========================================================
# CLEAN
# ==========================================================

def clean(t):

    return (
        re.sub(
            r"\s+",
            " ",
            t
        ).strip()
        if t
        else ""
    )


# ==========================================================
# HTTP
# ==========================================================

def get_soup(url):

    for attempt in range(1, 4):

        try:

            r = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            print(
                f"🌐 GET {r.status_code} | "
                f"{r.url}",
                flush=True
            )

            print(
                f"   HTML: {len(r.text)}",
                flush=True
            )

            # --------------------------------------------------
            # OK
            # --------------------------------------------------

            if r.status_code == 200:

                return BeautifulSoup(
                    r.text,
                    "html.parser"
                )

            # --------------------------------------------------
            # 429
            # --------------------------------------------------

            if r.status_code == 429:

                print(
                    "⚠️ HTTP 429",
                    flush=True
                )

                # На тесте НЕ долбим сайт много раз.
                # Если получили 429 — прекращаем запросы.

                return BeautifulSoup(
                    "",
                    "html.parser"
                )

            # --------------------------------------------------
            # OTHER ERROR
            # --------------------------------------------------

            print(
                f"⚠️ HTTP ERROR: "
                f"{r.status_code}",
                flush=True
            )

        except Exception as e:

            print(
                f"❌ GET ERROR: {e}",
                flush=True
            )

        if attempt < 3:

            time.sleep(2)

    return BeautifulSoup(
        "",
        "html.parser"
    )


# ==========================================================
# TEST MAIN PAGE
# ==========================================================

def test_main_page():

    print("")
    print("=" * 70)
    print("🧪 ТЕСТ ГЛАВНОЙ СТРАНИЦЫ")
    print("=" * 70)

    try:

        r = session.get(
            BASE,
            timeout=30,
            allow_redirects=True
        )

        print(
            f"STATUS: {r.status_code}",
            flush=True
        )

        print(
            f"URL: {r.url}",
            flush=True
        )

        print(
            f"HTML: {len(r.text)}",
            flush=True
        )

        print(
            f"TITLE: "
            f"{clean(BeautifulSoup(r.text, 'html.parser').title.get_text()) if BeautifulSoup(r.text, 'html.parser').title else ''}",
            flush=True
        )

        if r.status_code == 200:

            print(
                "✅ ГЛАВНАЯ ДОСТУПНА",
                flush=True
            )

            return True

        if r.status_code == 429:

            print(
                "❌ ГЛАВНАЯ ТОЖЕ 429",
                flush=True
            )

            return False

        print(
            f"❌ Главная вернула "
            f"HTTP {r.status_code}",
            flush=True
        )

        return False

    except Exception as e:

        print(
            f"❌ MAIN ERROR: {e}",
            flush=True
        )

        return False


# ==========================================================
# CATEGORIES
# ==========================================================

def get_categories():

    print("")
    print("=" * 70)
    print("🌳 ПОИСК ВСЕХ КАТЕГОРИЙ JMAX")
    print("=" * 70)

    soup = get_soup(
        BASE
    )

    if (
        not soup
        or
        not soup.find_all(True)
    ):

        print(
            "❌ Главная страница не загрузилась",
            flush=True
        )

        return []

    categories = []

    seen = set()

    # ======================================================
    # Ищем ссылки категорий
    # ======================================================

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = a.get(
            "href",
            ""
        ).strip()

        if not href:
            continue

        if not href.startswith(
            "http"
        ):

            href = (
                BASE +
                "/" +
                href.lstrip("/")
            )

        if (
            "route=product/category"
            not in href
        ):

            continue

        # Убираем page
        href = re.sub(
            r"[&?]page=\d+",
            "",
            href
        )

        if href in seen:
            continue

        seen.add(href)

        categories.append(
            href
        )

        print(
            f"📂 Категория "
            f"#{len(categories)}: "
            f"{href}",
            flush=True
        )

    print("")
    print(
        f"📂 ВСЕГО КАТЕГОРИЙ: "
        f"{len(categories)}",
        flush=True
    )

    return categories


# ==========================================================
# LAST PAGE
# ==========================================================

def get_last_page(soup):

    pages = [1]

    for a in soup.select(
        "ul.pagination a"
    ):

        href = a.get(
            "href",
            ""
        )

        m = re.search(
            r"[?&]page=(\d+)",
            href
        )

        if m:

            pages.append(
                int(m.group(1))
            )

    return max(pages)


# ==========================================================
# PARSE CATEGORY
# ==========================================================

def parse_category(
    cat_url
):

    all_items = []

    print("")
    print(
        f"📂 CATEGORY: "
        f"{cat_url}",
        flush=True
    )

    # ------------------------------------------------------
    # FIRST PAGE
    # ------------------------------------------------------

    first_page = get_soup(
        cat_url
    )

    if (
        not first_page
        or
        not first_page.find_all(True)
    ):

        print(
            "❌ Категория не загрузилась",
            flush=True
        )

        return []

    last_page = get_last_page(
        first_page
    )

    print(
        f"📄 Последняя страница: "
        f"{last_page}",
        flush=True
    )

    seen_products = set()

    # ======================================================
    # PAGES
    # ======================================================

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:

            url = cat_url

        else:

            url = (
                f"{cat_url}&page={page}"
            )

        print(
            f"📄 Страница "
            f"{page}/{last_page}",
            flush=True
        )

        soup = get_soup(
            url
        )

        if (
            not soup
            or
            not soup.find_all(True)
        ):

            print(
                "⛔ Страница не загрузилась",
                flush=True
            )

            break

        products = []

        # --------------------------------------------------
        # PRODUCT LINKS
        # --------------------------------------------------

        for a in soup.find_all(
            "a",
            href=True
        ):

            href = a.get(
                "href",
                ""
            ).strip()

            if not href:
                continue

            if not href.startswith(
                "http"
            ):

                href = (
                    BASE +
                    "/" +
                    href.lstrip("/")
                )

            if (
                "route=product/product"
                not in href
            ):

                continue

            href = href.split(
                "#"
            )[0]

            if href in products:
                continue

            products.append(
                href
            )

        # --------------------------------------------------
        # NO PRODUCTS
        # --------------------------------------------------

        if not products:

            print(
                "⛔ Товаров нет",
                flush=True
            )

            break

        # --------------------------------------------------
        # REMOVE DUPLICATES
        # --------------------------------------------------

        new_products = []

        for href in products:

            if href in seen_products:
                continue

            seen_products.add(
                href
            )

            new_products.append(
                href
            )

        if not new_products:

            print(
                "⛔ Новых товаров нет",
                flush=True
            )

            break

        print(
            f"📦 Товаров: "
            f"{len(new_products)}",
            flush=True
        )

        # --------------------------------------------------
        # PRODUCTS
        # --------------------------------------------------

        for href in new_products:

            item = parse_product(
                href
            )

            if item and item[1]:

                all_items.append(
                    item
                )

            time.sleep(
                0.1
            )

    print(
        f"✅ Категория закончена: "
        f"{len(all_items)} товаров",
        flush=True
    )

    return all_items


# ==========================================================
# PRODUCT SOUP
# ==========================================================

def parse_product_soup(
    soup,
    url
):

    # ======================================================
    # TITLE
    # ======================================================

    title = ""

    h1 = soup.select_one(
        "h1"
    )

    if h1:

        title = clean(
            h1.get_text()
        )

    # ======================================================
    # SKU
    # ======================================================

    sku = ""

    sku_tag = soup.select_one(
        ".product-data__item.model"
    )

    if sku_tag:

        sku = clean(
            sku_tag
            .get_text()
            .replace(
                "Код товара:",
                ""
            )
        )

    # ======================================================
    # PRICE
    # ======================================================

    price = ""

    p = soup.select_one(
        ".product-page__price"
    )

    if p:

        price = clean(
            p.get_text()
        )

    # ======================================================
    # STATUS
    # ======================================================

    status = ""

    btn = soup.select_one(
        "#button-cart span"
    )

    if btn:

        status = clean(
            btn.get_text()
        )

    else:

        text = clean(
            soup.get_text(
                " ",
                strip=True
            )
        ).lower()

        if (
            "нет в наличии"
            in text
        ):

            status = "Нет в наличии"

        else:

            status = "Нет кнопки"

    return [
        sku,
        title,
        price,
        status,
        url
    ]


# ==========================================================
# PRODUCT
# ==========================================================

def parse_product(
    url
):

    soup = get_soup(
        url
    )

    if not soup.select_one(
        "h1"
    ):

        return [
            "",
            "",
            "",
            "",
            url
        ]

    return parse_product_soup(
        soup,
        url
    )


# ==========================================================
# MAIN
# ==========================================================

def run_parser():

    if is_locked():

        print(
            "🔒 Парсер уже запущен",
            flush=True
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

        # ==================================================
        # 🧪 MAIN TEST
        # ==================================================

        if not test_main_page():

            print("")
            print(
                "=" * 70
            )

            print(
                "❌ RAILWAY НЕ ПУСКАЕТ JMAX",
                flush=True
            )

            print(
                "❌ Обычная главная страница "
                "тоже недоступна.",
                flush=True
            )

            print(
                "❌ Это не проблема логина.",
                flush=True
            )

            print(
                "=" * 70
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # ==================================================
        # CATEGORIES
        # ==================================================

        cats = get_categories()

        if CATEGORY_LIMIT:

            cats = cats[
                :CATEGORY_LIMIT
            ]

        total_categories = len(
            cats
        )

        if total_categories == 0:

            print(
                "❌ Категории не найдены",
                flush=True
            )

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            return

        # ==================================================
        # EXCEL
        # ==================================================

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        # ==================================================
        # PARSE
        # ==================================================

        seen = set()

        for i, cat in enumerate(
            cats,
            1
        ):

            progress = int(
                (i - 1)
                /
                total_categories
                *
                100
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )

            print("")
            print(
                "=" * 70
            )

            print(
                f"📂 КАТЕГОРИЯ "
                f"{i}/{total_categories}",
                flush=True
            )

            print(
                cat,
                flush=True
            )

            print(
                "=" * 70
            )

            items = parse_category(
                cat
            )

            print(
                f"📦 Получено: "
                f"{len(items)}",
                flush=True
            )

            # --------------------------------------------------
            # DEDUP
            # --------------------------------------------------

            for item in items:

                sku = item[0]
                title = item[1]
                price = item[2]
                status = item[3]
                url = item[4]

                if not title:
                    continue

                key = (
                    sku.strip()
                    if sku
                    else url.strip()
                )

                if not key:
                    continue

                if key in seen:
                    continue

                seen.add(
                    key
                )

                ws.append([
                    sku,
                    title,
                    price,
                    status,
                    url
                ])

            time.sleep(
                0.2
            )

        # ==================================================
        # SAVE
        # ==================================================

        print("")
        print(
            "=" * 70
        )

        print(
            "💾 Сохраняем Excel...",
            flush=True
        )

        tmp = (
            FILE_PATH +
            ".tmp"
        )

        wb.save(
            tmp
        )

        os.replace(
            tmp,
            FILE_PATH
        )

        # ==================================================
        # FINISH
        # ==================================================

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        print("")
        print(
            "=" * 70
        )

        print(
            "✅ ГОТОВО. "
            "Харьковская 4421-4422 Jmax",
            flush=True
        )

        print(
            f"📊 Всего товаров: "
            f"{len(seen)}",
            flush=True
        )

        print(
            f"📂 Категорий: "
            f"{total_categories}",
            flush=True
        )

        print(
            "=" * 70
        )

    finally:

        set_lock(False)


# ==========================================================
# START
# ==========================================================

if __name__ == "__main__":

    run_parser()
