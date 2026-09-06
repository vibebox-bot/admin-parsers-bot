import os
import json
import time
import requests
import xml.etree.ElementTree as ET
from datetime import datetime
from openpyxl import Workbook
import sys


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Maraton")


# =========================================================
# НАСТРОЙКИ
# =========================================================

XML_URL = "https://maraton.ua/yandexmarket/97211e2f-3247-455b-9cd3-59c441963309.xml"

CATEGORY_LIMIT = 2
# CATEGORY_LIMIT = None


OUTPUT_DIR = os.path.abspath("output/Maraton")
FILE_PATH = os.path.join(OUTPUT_DIR, "Maraton_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


HEADERS = {
    "User-Agent": "Mozilla/5.0"
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

    for _ in range(3):

        try:
            response = session.get(
                XML_URL,
                timeout=60
            )

            response.raise_for_status()

            return response.content

        except Exception:
            time.sleep(2)

    raise Exception(
        "Не удалось скачать XML Maraton"
    )


def get_text(element, tag):

    child = element.find(tag)

    if child is None:
        return ""

    return " ".join(
        "".join(child.itertext()).split()
    )


# =========================================================
# PARSE XML
# =========================================================

def parse_xml(xml_data):

    root = ET.fromstring(xml_data)

    offers = root.findall(".//offer")

    print(f"📦 Offers: {len(offers)}")

    return offers


# =========================================================
# RUN PARSER
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
        # DOWNLOAD XML
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

        seen = set()

        # -------------------------------------------------
        # PRODUCTS
        # -------------------------------------------------

        parse_total = len(offers_to_parse)

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

            # Артикул
            sku = offer.get(
                "id",
                ""
            ).strip()

            # URL
            url = get_text(
                offer,
                "url"
            )

            # Цена
            price = get_text(
                offer,
                "price"
            )

            # Валюта
            currency = get_text(
                offer,
                "currencyId"
            )

            # Категория
            category_id = get_text(
                offer,
                "categoryId"
            )

            # Название
            title = get_text(
                offer,
                "name"
            )

            # Наличие
            available = (
                offer.get(
                    "available",
                    ""
                )
                .strip()
                .lower()
            )

            if available == "true":
                status = "В наличии"
            else:
                status = "Нет в наличии"

            # -------------------------------------------------
            # DEDUP
            # -------------------------------------------------

            key = (
                sku
                or url
                or (title, price)
            )

            if key in seen:
                continue

            seen.add(key)

            if not title:
                continue

            # -------------------------------------------------
            # WRITE
            # -------------------------------------------------

            ws.append([
                sku,
                title,
                price,
                currency,
                status,
                category_id,
                url
            ])

        # -------------------------------------------------
        # SAVE EXCEL
        # -------------------------------------------------

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

        # -------------------------------------------------
        # FINISH
        # -------------------------------------------------

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
