import os
import json
import time
import requests
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import Workbook
from bs4 import BeautifulSoup
import sys


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Maraton")

# =========================================================
# XML
# =========================================================

XML_URL = "https://maraton.ua/yandexmarket/97211e2f-3247-455b-9cd3-59c441963309.xml"

# ТЕСТ
CATEGORY_LIMIT = 2

# ПОЛНЫЙ ЗАПУСК
# CATEGORY_LIMIT = None


# =========================================================
# PATHS
# =========================================================

OUTPUT_DIR = os.path.abspath("output/Maraton")
FILE_PATH = os.path.join(OUTPUT_DIR, "Maraton_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


# =========================================================
# SESSION
# =========================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
}

session = requests.Session()
session.headers.update(HEADERS)


# =========================================================
# LOCK
# =========================================================

def is_locked():
    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        if age > 3600:
            os.remove(LOCK_FILE)
            return False

        return True

    except Exception:
        return False


def set_lock(state):
    if state:
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)


# =========================================================
# STATUS
# =========================================================

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


# =========================================================
# XML
# =========================================================

def download_xml():

    for attempt in range(3):

        try:
            response = session.get(
                XML_URL,
                timeout=60
            )

            response.raise_for_status()

            return response.content

        except Exception as e:

            if attempt == 2:
                raise Exception(
                    f"Не удалось скачать XML Maraton: {e}"
                )

            time.sleep(2)


def get_text(element, tag):

    child = element.find(tag)

    if child is None:
        return ""

    return " ".join(
        "".join(child.itertext()).split()
    )


def parse_xml(xml_data):

    root = ET.fromstring(xml_data)

    offers = root.findall(".//offer")

    print(f"📦 Offers: {len(offers)}")

    return offers


# =========================================================
# PRODUCT PAGE
# =========================================================

def get_product_data(url, xml_price="", xml_available=""):

    if not url:
        return xml_price, (
            "В наличии"
            if xml_available == "true"
            else "Нет в наличии"
        )

    # XML иногда содержит http
    # переводим на https
    if url.startswith("http://maraton.ua"):
        url = url.replace(
            "http://maraton.ua",
            "https://maraton.ua",
            1
        )

    for attempt in range(3):

        try:

            response = session.get(
                url,
                timeout=30
            )

            response.raise_for_status()

            soup = BeautifulSoup(
                response.text,
                "html.parser"
            )

            # =================================================
            # ЦЕНА
            # =================================================

            price = ""

            # Основной вариант
            price_element = soup.select_one(
                ".price-number.s-product-price"
            )

            if price_element:

                price = (
                    price_element.get("data-price")
                    or price_element.get_text(
                        " ",
                        strip=True
                    )
                ).strip()

            # Запасной вариант:
            # <meta itemprop="price" content="39">
            if not price:

                meta_price = soup.select_one(
                    'meta[itemprop="price"]'
                )

                if meta_price:

                    price = (
                        meta_price.get("content")
                        or ""
                    ).strip()

            # =================================================
            # НАЛИЧИЕ
            # =================================================

            status = ""

            stock_element = soup.select_one(
                ".stock-text"
            )

            if stock_element:

                status = stock_element.get_text(
                    " ",
                    strip=True
                )

            # Если статус не нашли
            if not status:

                meta_availability = soup.select_one(
                    'link[itemprop="availability"]'
                )

                if meta_availability:

                    availability = (
                        meta_availability.get("href")
                        or ""
                    ).lower()

                    if "instock" in availability:
                        status = "В наличии"

                    elif "outofstock" in availability:
                        status = "Нет в наличии"

            # =================================================
            # FALLBACK
            # =================================================

            if not price:
                price = xml_price

            if not status:

                status = (
                    "В наличии"
                    if xml_available == "true"
                    else "Нет в наличии"
                )

            return price, status

        except Exception:

            if attempt < 2:
                time.sleep(1)
                continue

    # =====================================================
    # Если страница не открылась
    # =====================================================

    return (
        xml_price,
        (
            "В наличии"
            if xml_available == "true"
            else "Нет в наличии"
        )
    )


# =========================================================
# PARSER
# =========================================================

def run_parser():

    if is_locked():
        return

    set_lock(True)

    try:

        save_status(
            True,
            0,
            USER,
            FILE_PATH
        )

        # -------------------------------------------------
        # XML
        # -------------------------------------------------

        xml_data = download_xml()

        offers = parse_xml(xml_data)

        total = len(offers)

        if total == 0:

            save_status(
                False,
                100,
                USER,
                FILE_PATH
            )

            print("⚠ XML пустой")

            return

        # -------------------------------------------------
        # TEST MODE
        # -------------------------------------------------

        if CATEGORY_LIMIT:

            offers_to_parse = offers[:CATEGORY_LIMIT]

            print(
                f"🧪 TEST MODE: "
                f"{len(offers_to_parse)} offers"
            )

        else:

            offers_to_parse = offers

        # -------------------------------------------------
        # EXCEL
        # -------------------------------------------------

        wb = Workbook()

        ws = wb.active
        ws.title = "Maraton"

        ws.append([
            "Артикул",
            "Название",
            "Цена",
            "Валюта",
            "Наличие",
            "Категория",
            "URL"
        ])

        # -------------------------------------------------
        # DEDUP
        # -------------------------------------------------

        seen = set()

        # Чтобы не открывать один URL несколько раз
        price_cache = {}

        fallback_count = 0

        parse_total = len(offers_to_parse)

        # -------------------------------------------------
        # OFFERS
        # -------------------------------------------------

        for i, offer in enumerate(
            offers_to_parse,
            1
        ):

            progress = int(
                i / parse_total * 100
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )

            # ---------------------------------------------
            # XML DATA
            # ---------------------------------------------

            sku = (
                offer.get("id", "")
                .strip()
            )

            url = get_text(
                offer,
                "url"
            )

            xml_price = get_text(
                offer,
                "price"
            )

            currency = get_text(
                offer,
                "currencyId"
            )

            category_id = get_text(
                offer,
                "categoryId"
            )

            title = get_text(
                offer,
                "name"
            )

            available = (
                offer.get(
                    "available",
                    ""
                )
                .strip()
                .lower()
            )

            # ---------------------------------------------
            # DEDUP
            # ---------------------------------------------

            key = (
                sku
                or url
                or (
                    title,
                    xml_price
                )
            )

            if key in seen:
                continue

            seen.add(key)

            if not title:
                continue

            # ---------------------------------------------
            # PRODUCT PAGE
            # ---------------------------------------------

            if url in price_cache:

                price, status = price_cache[url]

            else:

                price, status = get_product_data(
                    url=url,
                    xml_price=xml_price,
                    xml_available=available
                )

                price_cache[url] = (
                    price,
                    status
                )

            # ---------------------------------------------
            # FALLBACK CHECK
            # ---------------------------------------------

            if price == xml_price:

                # Не обязательно значит ошибка:
                # цена сайта может совпасть с XML.
                #
                # Поэтому здесь ничего не считаем.
                pass

            # ---------------------------------------------
            # EXCEL
            # ---------------------------------------------

            ws.append([
                sku,
                title,
                price,
                currency,
                status,
                category_id,
                url
            ])

        # =================================================
        # SAVE
        # =================================================

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

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        print("✅ Готово. Maraton")

    finally:

        set_lock(False)


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    run_parser()
