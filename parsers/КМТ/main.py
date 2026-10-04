import os
import json
import time
import requests
import sys
import xml.etree.ElementTree as ET

from datetime import datetime
from openpyxl import Workbook


# ==========================================================
# 👤 USER
# ==========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"


print("🔥 Харьковская КМТ — XML PARSER")


# ==========================================================
# ⚙️ CONFIG
# ==========================================================

FEED_URL = "https://kmt5.com.ua/feed/alsj9tvf74xcmfavjl7rhkljz3os3kwy"


# None = все товары
# Например 100 = только первые 100 товаров для теста
PRODUCT_LIMIT = None
# PRODUCT_LIMIT = 100


OUTPUT_DIR = os.path.abspath("output/КМТ")

FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "КМТ_LIVE.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

LOCK_FILE = os.path.join(
    OUTPUT_DIR,
    "lock.txt"
)


HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 "
        "(Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 "
        "(KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    ),
    "Accept": (
        "application/xml,"
        "text/xml,"
        "application/xhtml+xml,"
        "text/html;q=0.9,"
        "*/*;q=0.8"
    )
}


session = requests.Session()

session.headers.update(
    HEADERS
)


# ==========================================================
# 🔒 LOCK
# ==========================================================

def is_locked():

    if not os.path.exists(LOCK_FILE):

        return False


    try:

        age = (
            time.time()
            - os.path.getmtime(LOCK_FILE)
        )


        # Если lock старше часа —
        # считаем его зависшим

        if age > 3600:

            os.remove(
                LOCK_FILE
            )

            return False


        return True


    except Exception:

        return False


# ==========================================================
# 🔒 SET LOCK
# ==========================================================

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

            try:

                os.remove(
                    LOCK_FILE
                )

            except Exception:

                pass


# ==========================================================
# 📊 STATUS
# ==========================================================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path="",
    error=False,
    error_message=""
):

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )


    data = {

        "running": running,

        "progress": progress,

        "user": user,

        "time": datetime.now().strftime(
            "%Y-%m-%d %H:%M:%S"
        ),

        "file_path": file_path,

        "error": error,

        "error_message": error_message
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


# ==========================================================
# 🧹 CLEAN TEXT
# ==========================================================

def clean(value):

    if value is None:

        return ""


    return " ".join(
        str(value).split()
    ).strip()


# ==========================================================
# 🌐 DOWNLOAD XML
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

                last_error = (
                    "XML feed пустой"
                )

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
        f"Не удалось загрузить XML: "
        f"{last_error}"
    )


# ==========================================================
# 📦 XML PARSE
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


    # ------------------------------------------------------
    # Ищем offer независимо от структуры YML
    # ------------------------------------------------------

    offers = root.findall(
        ".//offer"
    )


    if not offers:

        raise RuntimeError(
            "В XML отсутствуют элементы <offer>"
        )


    print(
        f"📦 Найдено товаров в XML: "
        f"{len(offers)}"
    )


    if PRODUCT_LIMIT:

        offers = offers[
            :PRODUCT_LIMIT
        ]

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

        vendor_code = clean(
            offer.findtext(
                "vendorCode",
                default=""
            )
        )


        # ==================================================
        # PRODUCT CODE
        # ==================================================

        product_code = ""


        for param in offer.findall(
            "param"
        ):

            name = clean(
                param.get(
                    "name",
                    ""
                )
            )


            if name.lower() == "код":

                product_code = clean(
                    param.text
                )

                break


        # ==================================================
        # TITLE
        # ==================================================

        title = clean(
            offer.findtext(
                "name",
                default=""
            )
        )


        # ==================================================
        # PRICE
        # ==================================================

        price = clean(
            offer.findtext(
                "price",
                default=""
            )
        )


        # ==================================================
        # STATUS
        # ==================================================

        available = clean(
            offer.get(
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


        # ==================================================
        # URL
        # ==================================================

        url = clean(
            offer.findtext(
                "url",
                default=""
            )
        )


        # ==================================================
        # ПРОВЕРКА
        # ==================================================

        # Если товар вообще не имеет названия,
        # не добавляем его в результат.

        if not title:

            continue


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
            or index % 100 == 0
            or index == total
        ):

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )


            print(
                f"📦 {index}/{total} "
                f"({progress}%)"
            )


    if not result:

        raise RuntimeError(
            "После обработки XML "
            "не получено ни одного товара"
        )


    return result


# ==========================================================
# 📄 CREATE EMPTY EXCEL
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


    wb.save(
        tmp
    )


    os.replace(
        tmp,
        FILE_PATH
    )


    print(
        "⚠ КМТ_LIVE.xlsx создан пустым"
    )


# ==========================================================
# 📄 SAVE EXCEL
# ==========================================================

def save_excel(items):

    if not items:

        raise RuntimeError(
            "Попытка сохранить пустой результат"
        )


    wb = Workbook()

    ws = wb.active

    ws.title = "КМТ"


    # ======================================================
    # HEADERS
    # ======================================================

    ws.append([
        "CODE",
        "PRODUCT_CODE",
        "TITLE",
        "PRICE",
        "STATUS",
        "URL"
    ])


    # ======================================================
    # DATA
    # ======================================================

    for item in items:

        ws.append(
            item
        )


    # ======================================================
    # SAVE
    # ======================================================

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
        f"💾 Сохранено товаров: "
        f"{len(items)}"
    )


# ==========================================================
# 🚀 MAIN
# ==========================================================

def run_parser():

    if is_locked():

        print(
            "⚠ Парсер КМТ уже запущен"
        )

        return


    set_lock(True)


    try:

        # ==================================================
        # START STATUS
        # ==================================================

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
        # DOWNLOAD
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
        # FINAL CHECK
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

            save_excel(
                items
            )


        except Exception as e:

            error_message = (
                f"Ошибка сохранения Excel: {e}"
            )


            print(
                f"❌ ОШИБКА КМТ: "
                f"{error_message}"
            )


            # При ошибке сохранения тоже
            # создаём пустой Excel

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
        # SUCCESS
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
# START
# ==========================================================

if __name__ == "__main__":

    run_parser()
