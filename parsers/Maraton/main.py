import os
import sys
import json
import time
import re

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from urllib.parse import urlparse, parse_qs


# =========================================================
# НАСТРОЙКИ
# =========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

XML_URL = (
    "https://maraton.ua/yandexmarket/"
    "97211e2f-3247-455b-9cd3-59c441963309.xml"
)

OUTPUT_DIR = os.path.abspath("output/Maraton")
FILE_PATH = os.path.join(OUTPUT_DIR, "Maraton_LIVE.xlsx")

STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "parser.lock")

# Для теста — первые 2 товара
# После проверки поставь None
CATEGORY_LIMIT = 25


# =========================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# =========================================================

def clean_text(value):
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value).replace("\xa0", " ")
    ).strip()


def write_status(progress=0, running=True):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": USER,
        "time": time.time(),
        "file_path": FILE_PATH,
    }

    try:
        with open(
            STATUS_PATH,
            "w",
            encoding="utf-8"
        ) as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )
    except Exception:
        pass


def acquire_lock():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if os.path.exists(LOCK_FILE):
        try:
            age = time.time() - os.path.getmtime(LOCK_FILE)

            if age > 3600:
                os.remove(LOCK_FILE)
            else:
                print("⚠ Maraton уже запущен")
                return False

        except Exception:
            pass

    try:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(os.getpid()))

        return True

    except Exception:
        return False


def release_lock():
    try:
        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
    except Exception:
        pass


# =========================================================
# SESSION
# =========================================================

def create_session():
    session = requests.Session()

    retry = Retry(
        total=3,
        connect=3,
        read=3,
        backoff_factor=1,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"],
    )

    adapter = HTTPAdapter(
        max_retries=retry,
        pool_connections=10,
        pool_maxsize=10,
    )

    session.mount("http://", adapter)
    session.mount("https://", adapter)

    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 "
            "(Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/140.0 Safari/537.36"
        ),
        "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
    })

    return session


# =========================================================
# XML
# =========================================================

def parse_xml(session):
    response = session.get(
        XML_URL,
        timeout=60
    )

    response.raise_for_status()

    soup = BeautifulSoup(
        response.content,
        "xml"
    )

    offers = []

    for offer in soup.find_all("offer"):

        url_el = offer.find("url")
        name_el = offer.find("name")

        if not url_el or not name_el:
            continue

        url = clean_text(
            url_el.get_text(" ", strip=True)
        )

        name = clean_text(
            name_el.get_text(" ", strip=True)
        )

        if not url:
            continue

        if not name:
            name = url

        offers.append({
            "name": name,
            "url": url,
        })

    return offers


# =========================================================
# ПАРСИНГ СТРАНИЦЫ ТОВАРА
# =========================================================

def parse_product_page(session, url):

    response = session.get(
        url,
        timeout=30,
        allow_redirects=True
    )

    response.raise_for_status()

    soup = BeautifulSoup(
        response.text,
        "html.parser"
    )

    # -----------------------------------------------------
    # SKU из URL
    #
    # Например:
    # ?sku=4562
    # -----------------------------------------------------

    parsed_url = urlparse(url)

    query = parse_qs(
        parsed_url.query
    )

    sku = query.get(
        "sku",
        [""]
    )[0].strip()

    price = ""
    currency = "USD"
    stock = ""

    # -----------------------------------------------------
    # Ищем input нужного SKU
    # -----------------------------------------------------

    sku_input = None

    if sku:

        for inp in soup.select(
            'input[name="sku_id"]'
        ):

            value = clean_text(
                inp.get("value")
            )

            if value == sku:
                sku_input = inp
                break

    # -----------------------------------------------------
    # Если SKU в URL нет —
    # берём выбранный вариант
    # -----------------------------------------------------

    if not sku_input:

        sku_input = soup.select_one(
            'input[name="sku_id"][checked]'
        )

    # =====================================================
    # ЕСЛИ НАШЛИ SKU
    # =====================================================

    if sku_input:

        # -------------------------------------------------
        # ЦЕНА
        # -------------------------------------------------

        price = clean_text(
            sku_input.get("data-price")
        )

        # -------------------------------------------------
        # LABEL ВАРИАНТА
        # -------------------------------------------------

        label = sku_input.find_parent("label")

        if label:

            # ---------------------------------------------
            # ВАЛЮТА
            # ---------------------------------------------

            currency_meta = label.select_one(
                'meta[itemprop="priceCurrency"]'
            )

            if currency_meta:

                currency = clean_text(
                    currency_meta.get("content")
                )

                if not currency:
                    currency = "USD"

        # -------------------------------------------------
        # НАЛИЧИЕ ПО SKU
        #
        # Например:
        # .sku-4562-stock
        # -------------------------------------------------

        if sku:

            stock_wrapper = soup.select_one(
                f".sidebar__stock-wrapper.sku-{sku}-stock"
            )

            if stock_wrapper:

                stock_el = stock_wrapper.select_one(
                    ".stock-text"
                )

                if stock_el:

                    stock = clean_text(
                        stock_el.get_text(
                            " ",
                            strip=True
                        )
                    )

                # Если stock-text почему-то нет
                if not stock:

                    stock = clean_text(
                        stock_wrapper.get_text(
                            " ",
                            strip=True
                        )
                    )

        # -------------------------------------------------
        # Запасной вариант наличия
        # из LABEL
        # -------------------------------------------------

        if not stock and label:

            availability = label.select_one(
                'link[itemprop="availability"]'
            )

            if availability:

                href = clean_text(
                    availability.get("href")
                ).lower()

                if "instock" in href:
                    stock = "В наличии"

                elif "outofstock" in href:
                    stock = "Нет в наличии"

    # =====================================================
    # FALLBACK ЦЕНЫ
    # =====================================================

    if not price:

        price_el = soup.select_one(
            ".item-sidebar__price "
            ".price-number.s-product-price"
        )

        if price_el:

            # Берём именно ВИДИМЫЙ текст:
            # $72
            #
            # а НЕ data-price="1.61"
            #

            price_text = clean_text(
                price_el.get_text(
                    " ",
                    strip=True
                )
            )

            price_text = (
                price_text
                .replace("$", "")
                .replace(",", ".")
                .strip()
            )

            price = price_text

    # =====================================================
    # FALLBACK НАЛИЧИЯ
    # =====================================================

    if not stock:

        for el in soup.select(
            ".sidebar__stock-wrapper"
        ):

            style = (
                el.get("style") or ""
            ).lower()

            # Пропускаем скрытые варианты
            if "display: none" in style:
                continue

            stock_el = el.select_one(
                ".stock-text"
            )

            if stock_el:

                stock = clean_text(
                    stock_el.get_text(
                        " ",
                        strip=True
                    )
                )

                if stock:
                    break

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
        exist_ok=True
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

        ws.append([
            row["name"],
            row["price"],
            row["currency"],
            row["stock"],
            row["url"],
        ])

    # Ширина колонок
    ws.column_dimensions["A"].width = 65
    ws.column_dimensions["B"].width = 15
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 20
    ws.column_dimensions["E"].width = 80

    # Временный файл
    temp_file = FILE_PATH + ".tmp"

    wb.save(temp_file)

    os.replace(
        temp_file,
        FILE_PATH
    )


# =========================================================
# MAIN
# =========================================================

def run_parser():

    print("🔥 Maraton")

    if not acquire_lock():
        return

    write_status(
        progress=0,
        running=True
    )

    try:

        session = create_session()

        # -------------------------------------------------
        # XML
        # -------------------------------------------------

        offers = parse_xml(session)

        print(
            f"📦 Offers: {len(offers)}"
        )

        # -------------------------------------------------
        # TEST MODE
        # -------------------------------------------------

        if CATEGORY_LIMIT is not None:

            offers = offers[
                :CATEGORY_LIMIT
            ]

            print(
                f"🧪 TEST MODE: {len(offers)} offers"
            )

        total = len(offers)

        rows = []

        price_ok = 0
        price_errors = 0

        # -------------------------------------------------
        # ТОВАРЫ
        # -------------------------------------------------

        for index, offer in enumerate(
            offers,
            start=1
        ):

            url = offer["url"]

            try:

                page_data = parse_product_page(
                    session,
                    url
                )

                price = page_data["price"]

                if price:
                    price_ok += 1
                else:
                    price_errors += 1

                rows.append({
                    "name": offer["name"],
                    "price": price,
                    "currency": page_data["currency"],
                    "stock": page_data["stock"],

                    # ВАЖНО:
                    # только URL из XML
                    "url": offer["url"],
                })

            except Exception as e:

                price_errors += 1

                print(
                    f"⚠ Ошибка: {url} | {e}"
                )

                rows.append({
                    "name": offer["name"],
                    "price": "",
                    "currency": "USD",
                    "stock": "",
                    "url": offer["url"],
                })

            # -------------------------------------------------
            # ПРОГРЕСС
            # -------------------------------------------------

            if total:

                progress = int(
                    index / total * 100
                )

                write_status(
                    progress=progress,
                    running=True
                )

        # -------------------------------------------------
        # СОХРАНЕНИЕ
        # -------------------------------------------------

        save_excel(rows)

        print(
            f"💰 Prices from site: "
            f"{price_ok}/{total}"
        )

        print(
            f"⚠ Price errors: "
            f"{price_errors}"
        )

        print(
            f"📦 Products: "
            f"{len(rows)}"
        )

        print(
            "✅ Готово. Maraton"
        )

        write_status(
            progress=100,
            running=False
        )

    except Exception as e:

        print(
            f"❌ Maraton ERROR: {e}"
        )

        write_status(
            progress=0,
            running=False
        )

        raise

    finally:

        release_lock()


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    run_parser()
