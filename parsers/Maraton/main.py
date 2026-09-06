import os
import sys
import json
import time
import requests
import xml.etree.ElementTree as ET

from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry


# =========================================================
# SETTINGS
# =========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

XML_URL = (
    "https://maraton.ua/yandexmarket/"
    "97211e2f-3247-455b-9cd3-59c441963309.xml"
)

OUTPUT_DIR = os.path.abspath("output/Maraton")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Maraton_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "lock.txt"
)


# =========================================================
# TEST MODE
# =========================================================

CATEGORY_LIMIT = 25

# Для полного запуска:
# CATEGORY_LIMIT = None


# =========================================================
# HTTP
# =========================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,"
        "application/xml;q=0.9,image/avif,"
        "image/webp,*/*;q=0.8"
    ),
    "Accept-Language": (
        "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7"
    ),
    "Connection": "keep-alive",
}


def get_session():
    session = requests.Session()

    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=0.5,
        status_forcelist=[
            429,
            500,
            502,
            503,
            504,
        ],
        allowed_methods=[
            "GET",
        ],
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=10,
        pool_maxsize=10,
    )

    session.mount(
        "https://",
        adapter,
    )

    session.mount(
        "http://",
        adapter,
    )

    session.headers.update(HEADERS)

    return session


# =========================================================
# HELPERS
# =========================================================

def clean_text(value):
    if not value:
        return ""

    return " ".join(
        str(value).split()
    ).strip()


def normalize_url(url):
    url = clean_text(url)

    if url.startswith("//"):
        return "https:" + url

    if url.startswith("http://"):
        return "https://" + url[7:]

    return url


# =========================================================
# STATUS
# =========================================================

def write_status(
    running=True,
    progress=0,
):
    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True,
    )

    data = {
        "running": running,
        "progress": progress,
        "user": USER,
        "time": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),
        "file_path": FILE_PATH,
    }

    tmp_path = STATUS_PATH + ".tmp"

    with open(
        tmp_path,
        "w",
        encoding="utf-8",
    ) as f:
        json.dump(
            data,
            f,
            ensure_ascii=False,
            indent=2,
        )

    os.replace(
        tmp_path,
        STATUS_PATH,
    )


# =========================================================
# LOCK
# =========================================================

def acquire_lock():
    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True,
    )

    if os.path.exists(LOCK_FILE):
        try:
            age = (
                time.time()
                - os.path.getmtime(LOCK_FILE)
            )

            if age > 3600:
                os.remove(LOCK_FILE)
            else:
                raise RuntimeError(
                    "Maraton parser уже запущен."
                )

        except FileNotFoundError:
            pass

    with open(
        LOCK_FILE,
        "w",
        encoding="utf-8",
    ) as f:
        f.write(
            json.dumps(
                {
                    "user": USER,
                    "time": datetime.now().strftime(
                        "%Y-%m-%d %H:%M:%S"
                    ),
                },
                ensure_ascii=False,
            )
        )


def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


# =========================================================
# XML
# =========================================================

def fetch_xml(session):
    response = session.get(
        XML_URL,
        timeout=60,
    )

    response.raise_for_status()

    return response.content


def parse_xml(xml_data):
    root = ET.fromstring(xml_data)

    offers = []

    for offer in root.findall(".//offer"):

        name = clean_text(
            offer.findtext("name")
        )

        url = clean_text(
            offer.findtext("url")
        )

        if not name or not url:
            continue

        offers.append(
            {
                "name": name,
                "url": normalize_url(url),
            }
        )

    return offers


# =========================================================
# PRODUCT PAGE
# =========================================================

def parse_product_page(
    session,
    url,
):
    response = session.get(
        url,
        timeout=30,
        allow_redirects=True,
    )

    response.raise_for_status()

    soup = BeautifulSoup(
        response.text,
        "html.parser",
    )

    # -----------------------------------------------------
    # PRICE
    # -----------------------------------------------------

    price = ""

    # Основной селектор из карточки товара.
    price_el = soup.select_one(
        ".item-sidebar__price "
        ".price-number.s-product-price"
    )

    if price_el:

        price = clean_text(
            price_el.get("data-price")
        )

        if not price:
            price = clean_text(
                price_el.get_text(
                    " ",
                    strip=True,
                )
            )

    # Если основной селектор не найден.
    if not price:

        price_el = soup.select_one(
            ".price-number.s-product-price"
        )

        if price_el:

            price = clean_text(
                price_el.get("data-price")
            )

            if not price:
                price = clean_text(
                    price_el.get_text(
                        " ",
                        strip=True,
                    )
                )

    # -----------------------------------------------------
    # CURRENCY
    # -----------------------------------------------------

    currency = ""

    currency_el = soup.select_one(
        'meta[itemprop="priceCurrency"]'
    )

    if currency_el:
        currency = clean_text(
            currency_el.get("content")
        )

    if not currency:
        currency = "USD"

    # -----------------------------------------------------
    # STOCK
    # -----------------------------------------------------

    stock = ""

    stock_el = soup.select_one(
        ".stock-text"
    )

    if stock_el:

        stock = clean_text(
            stock_el.get_text(
                " ",
                strip=True,
            )
        )

    # Schema.org как запасной вариант.
    if not stock:

        availability_el = soup.select_one(
            'link[itemprop="availability"]'
        )

        if availability_el:

            href = clean_text(
                availability_el.get("href")
            ).lower()

            if "instock" in href:
                stock = "В наличии"

            elif "outofstock" in href:
                stock = "Нет в наличии"

            elif "preorder" in href:
                stock = "Под заказ"

    return {
        "price": price,
        "currency": currency,
        "stock": stock,
    }


# =========================================================
# EXCEL
# =========================================================

def save_excel(rows):
    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True,
    )

    wb = Workbook()

    ws = wb.active
    ws.title = "Maraton"

    headers = [
        "Название",
        "Цена",
        "Валюта",
        "Наличие",
        "Ссылка",
    ]

    ws.append(headers)

    for row in rows:

        ws.append(
            [
                row["name"],
                row["price"],
                row["currency"],
                row["stock"],
                row["url"],
            ]
        )

    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 90

    # Цена числом
    for cell in ws["B"][1:]:

        if cell.value:

            try:
                cell.value = float(
                    str(cell.value)
                    .replace(",", ".")
                    .replace("$", "")
                    .strip()
                )

            except Exception:
                pass

    tmp_file = FILE_PATH + ".tmp"

    wb.save(tmp_file)

    os.replace(
        tmp_file,
        FILE_PATH,
    )


# =========================================================
# RUN
# =========================================================

def run_parser():

    print("🔥 Maraton")

    acquire_lock()

    write_status(
        True,
        0,
    )

    try:

        session = get_session()

        # -------------------------------------------------
        # XML
        # -------------------------------------------------

        xml_data = fetch_xml(
            session
        )

        offers = parse_xml(
            xml_data
        )

        print(
            f"📦 Offers: {len(offers)}"
        )

        if CATEGORY_LIMIT is not None:

            offers = offers[
                :CATEGORY_LIMIT
            ]

            print(
                f"🧪 TEST MODE: "
                f"{len(offers)} offers"
            )

        # -------------------------------------------------
        # PRODUCTS
        # -------------------------------------------------

        rows = []

        prices_ok = 0
        price_errors = 0

        total = len(offers)

        for index, offer in enumerate(
            offers,
            start=1,
        ):

            price = ""
            currency = ""
            stock = ""

            try:

                result = parse_product_page(
                    session,
                    offer["url"],
                )

                price = result["price"]
                currency = result["currency"]
                stock = result["stock"]

            except Exception:

                price_errors += 1

            if price:
                prices_ok += 1

            rows.append(
                {
                    "name": offer["name"],
                    "price": price,
                    "currency": currency,
                    "stock": stock,
                    "url": offer["url"],
                }
            )

            progress = (
                int(index / total * 100)
                if total
                else 100
            )

            write_status(
                True,
                progress,
            )

            if index < total:
                time.sleep(0.15)

        # -------------------------------------------------
        # LOG
        # -------------------------------------------------

        print(
            f"💰 Prices from site: "
            f"{prices_ok}/{total}"
        )

        if price_errors:
            print(
                f"⚠ Price errors: "
                f"{price_errors}"
            )

        # -------------------------------------------------
        # SAVE
        # -------------------------------------------------

        save_excel(rows)

        print(
            f"📦 Products: {len(rows)}"
        )

        write_status(
            False,
            100,
        )

        print(
            "✅ Готово. Maraton"
        )

    except Exception as e:

        write_status(
            False,
            0,
        )

        print(
            f"❌ Ошибка Maraton: {e}"
        )

        raise

    finally:

        release_lock()


if __name__ == "__main__":
    run_parser()
