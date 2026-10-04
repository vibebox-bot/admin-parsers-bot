import os
import json
import time
import requests
import sys
import xml.etree.ElementTree as ET

from datetime import datetime
from openpyxl import Workbook
from bs4 import BeautifulSoup

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Харьковская КМТ — XML + SITE PRICE PARSER")


# ==========================================================
# НАСТРОЙКИ
# ==========================================================

FEED_URL = "https://kmt5.com.ua/feed/alsj9tvf74xcmfavjl7rhkljz3os3kwy"

PRODUCT_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/КМТ")
FILE_PATH = os.path.join(OUTPUT_DIR, "КМТ_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,"
        "application/xhtml+xml,"
        "application/xml,"
        "text/xml,"
        "*/*;q=0.8"
    )
}


session = requests.Session()
session.headers.update(HEADERS)


# ==========================================================
# LOCK
# ==========================================================

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

            try:
                os.remove(LOCK_FILE)
            except Exception:
                pass


# ==========================================================
# STATUS
# ==========================================================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path="",
    error=False,
    error_message=""
):

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path,
        "error": error,
        "error_message": error_message
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


# ==========================================================
# CLEAN
# ==========================================================

def clean(value):

    if value is None:
        return ""

    return " ".join(
        str(value).split()
    ).strip()


# ==========================================================
# XML
# ==========================================================

def download_feed():

    last_error = ""

    for attempt in range(3):

        try:

            print(
                f"🌐 Загрузка XML... "
                f"попытка {attempt + 1}/3"
            )

            response = session.get(
                FEED_URL,
                timeout=60
            )

            print(
                f"   HTTP: {response.status_code}"
            )

            if response.status_code != 200:

                last_error = (
                    f"HTTP {response.status_code}"
                )

                time.sleep(2)
                continue

            content = response.content

            if not content:

                last_error = "XML feed пустой"

                time.sleep(2)
                continue

            print(
                f"   Получено: "
                f"{len(content):,} байт"
            )

            return content

        except Exception as e:

            last_error = str(e)

            print(
                f"⚠ Ошибка загрузки: {e}"
            )

            time.sleep(2)

    raise RuntimeError(
        f"Не удалось загрузить XML: {last_error}"
    )


# ==========================================================
# XML HELPERS
# ==========================================================

def local_name(tag):

    if not isinstance(tag, str):
        return ""

    if "}" in tag:
        return tag.split("}", 1)[1]

    return tag


def get_child_text(element, wanted_name):

    wanted_name = wanted_name.lower()

    for child in element:

        if local_name(child.tag).lower() == wanted_name:

            return clean(child.text)

    return ""


def get_params(offer):

    result = {}

    for child in offer:

        if local_name(child.tag).lower() != "param":
            continue

        name = clean(
            child.attrib.get("name", "")
        )

        value = clean(child.text)

        if name:
            result[name.lower()] = value

    return result


# ==========================================================
# PRICE FROM PRODUCT PAGE
# ==========================================================

def get_site_price(url):

    if not url:
        return ""

    try:

        response = session.get(
            url,
            timeout=30
        )

        if response.status_code != 200:

            print(
                f"   ⚠ Цена: HTTP "
                f"{response.status_code}"
            )

            return ""

        soup = BeautifulSoup(
            response.content,
            "html.parser"
        )

        # ==================================================
        # ОСНОВНАЯ ЦЕНА КМТ
        #
        # <span class="opt"
        #       data-pdprice=""
        #       data-baseprice="1.10">
        #       $1.10
        # </span>
        #
        # Берём именно data-baseprice
        # ==================================================

        price_element = soup.select_one(
            ".bb-price .opt[data-baseprice]"
        )

        if price_element:

            price = clean(
                price_element.get(
                    "data-baseprice",
                    ""
                )
            )

            if price:
                return price

        # ==================================================
        # ДОПОЛНИТЕЛЬНЫЙ ПОИСК
        # если немного изменится HTML
        # ==================================================

        price_element = soup.select_one(
            ".opt[data-baseprice]"
        )

        if price_element:

            price = clean(
                price_element.get(
                    "data-baseprice",
                    ""
                )
            )

            if price:
                return price

        return ""

    except Exception as e:

        print(
            f"   ⚠ Ошибка получения цены: {e}"
        )

        return ""


# ==========================================================
# PARSE XML
# ==========================================================

def parse_feed(xml_content):

    try:

        root = ET.fromstring(
            xml_content
        )

    except ET.ParseError as e:

        raise RuntimeError(
            f"XML повреждён: {e}"
        )

    except Exception as e:

        raise RuntimeError(
            f"Ошибка чтения XML: {e}"
        )


    # Ищем offer независимо от namespace

    offers = [
        element
        for element in root.iter()
        if local_name(element.tag).lower() == "offer"
    ]


    if not offers:

        raise RuntimeError(
            "В XML отсутствуют элементы <offer>"
        )


    print(
        f"📦 Найдено товаров в XML: "
        f"{len(offers)}"
    )


    if PRODUCT_LIMIT:

        offers = offers[:PRODUCT_LIMIT]

        print(
            f"🧪 Тестовый лимит: "
            f"{PRODUCT_LIMIT}"
        )


    result = []

    total = len(offers)


    for index, offer in enumerate(
        offers,
        1
    ):

        # ==================================================
        # CODE
        # ==================================================

        vendor_code = get_child_text(
            offer,
            "vendorCode"
        )


        # ==================================================
        # PRODUCT_CODE
        # <param name="Код">792278</param>
        # ==================================================

        params = get_params(offer)

        product_code = params.get(
            "код",
            ""
        )


        # ==================================================
        # TITLE
        # ==================================================

        title = get_child_text(
            offer,
            "name"
        )


        # ==================================================
        # URL
        # ==================================================

        url = get_child_text(
            offer,
            "url"
        )


        # ==================================================
        # STATUS
        # ==================================================

        available = clean(
            offer.attrib.get(
                "available",
                ""
            )
        ).lower()


        if available == "true":

            status = "В наявності"

        elif available == "false":

            status = "Немає в наявності"

        else:

            status = available


        if not title:

            continue


        # ==================================================
        # PRICE FROM SITE
        # ==================================================

        print(
            f"💰 {index}/{total} "
            f"| {title[:60]}"
        )

        price = get_site_price(
            url
        )


        if price:

            print(
                f"   💵 Цена КМТ: ${price}"
            )

        else:

            print(
                "   ⚠ Цена на странице не найдена"
            )


        result.append([
            vendor_code,
            product_code,
            title,
            price,
            status,
            url
        ])


        # ==================================================
        # PROGRESS
        # ==================================================

        progress = int(
            index / total * 100
        )


        if (
            index == 1
            or index % 10 == 0
            or index == total
        ):

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )


    if not result:

        raise RuntimeError(
            "После обработки XML "
            "не получено ни одного товара"
        )


    return result


# ==========================================================
# EMPTY EXCEL
# ==========================================================

def create_empty_excel():

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )


    wb = Workbook()

    ws = wb.active

    ws.title = "КМТ"


    ws.append([
        "CODE",
        "PRODUCT_CODE",
        "TITLE",
        "PRICE",
        "STATUS",
        "URL"
    ])


    tmp = FILE_PATH + ".tmp"

    wb.save(tmp)

    os.replace(
        tmp,
        FILE_PATH
    )


    print(
        "⚠ КМТ_LIVE.xlsx создан пустым"
    )


# ==========================================================
# SAVE EXCEL
# ==========================================================

def save_excel(items):

    if not items:

        raise RuntimeError(
            "Попытка сохранить пустой результат"
        )


    wb = Workbook()

    ws = wb.active

    ws.title = "КМТ"


    ws.append([
        "CODE",
        "PRODUCT_CODE",
        "TITLE",
        "PRICE",
        "STATUS",
        "URL"
    ])


    for item in items:

        ws.append(item)


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


    print(
        f"💾 Сохранено товаров: "
        f"{len(items)}"
    )


# ==========================================================
# RUN
# ==========================================================

def run_parser():

    if is_locked():

        print(
            "⚠ Парсер КМТ уже запущен"
        )

        return


    set_lock(True)


    try:

        save_status(
            True,
            0,
            USER,
            FILE_PATH,
            False,
            ""
        )


        print(
            "🚀 Запуск парсера КМТ"
        )


        # ==================================================
        # XML
        # ==================================================

        try:

            xml_content = download_feed()

        except Exception as e:

            error_message = str(e)

            print(
                f"❌ ОШИБКА КМТ: "
                f"{error_message}"
            )


            create_empty_excel()


            save_status(
                False,
                100,
                USER,
                FILE_PATH,
                True,
                error_message
            )


            return


        # ==================================================
        # PARSE
        # ==================================================

        try:

            items = parse_feed(
                xml_content
            )

        except Exception as e:

            error_message = str(e)

            print(
                f"❌ ОШИБКА КМТ: "
                f"{error_message}"
            )


            create_empty_excel()


            save_status(
                False,
                100,
                USER,
                FILE_PATH,
                True,
                error_message
            )


            return


        # ==================================================
        # EMPTY RESULT
        # ==================================================

        if not items:

            error_message = (
                "Парсер завершился: "
                "товаров получено 0"
            )


            print(
                f"❌ ОШИБКА КМТ: "
                f"{error_message}"
            )


            create_empty_excel()


            save_status(
                False,
                100,
                USER,
                FILE_PATH,
                True,
                error_message
            )


            return


        # ==================================================
        # SAVE
        # ==================================================

        try:

            save_excel(items)

        except Exception as e:

            error_message = (
                f"Ошибка сохранения Excel: {e}"
            )


            print(
                f"❌ ОШИБКА КМТ: "
                f"{error_message}"
            )


            try:

                create_empty_excel()

            except Exception:

                pass


            save_status(
                False,
                100,
                USER,
                FILE_PATH,
                True,
                error_message
            )


            return


        # ==================================================
        # DONE
        # ==================================================

        save_status(
            False,
            100,
            USER,
            FILE_PATH,
            False,
            ""
        )


        print(
            "========================================"
        )

        print(
            "✅ КМТ ГОТОВО"
        )

        print(
            f"✅ Товаров: {len(items)}"
        )

        print(
            f"📄 Файл: {FILE_PATH}"
        )

        print(
            "========================================"
        )


    finally:

        set_lock(False)


# ==========================================================
# MAIN
# ==========================================================

if __name__ == "__main__":

    run_parser()
