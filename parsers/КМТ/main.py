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


print("🔥 Харьковская КМТ")

# ==========================================================
# НАСТРОЙКИ
# ==========================================================

FEED_URL = (
    "https://kmt5.com.ua/feed/"
    "alsj9tvf74xcmfavjl7rhkljz3os3kwy"
)

LOGIN_URL = "https://kmt5.com.ua/login/"
LOGIN_AJAX_URL = "https://kmt5.com.ua/login/?ajax=1"

# ==========================================================
# ЛОГИН КМТ
# ==========================================================

KMT_LOGIN = "finik257@gmail.com"
KMT_PASSWORD = "18022021"

# ==========================================================
# ТЕСТ
# После проверки поставить None
# ==========================================================

PRODUCT_LIMIT = None


# ==========================================================
# ФАЙЛЫ
# ==========================================================

OUTPUT_DIR = os.path.abspath(
    "output/КМТ"
)

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


# ==========================================================
# HEADERS
# ==========================================================

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
    ),

    "Accept-Language": (
        "uk-UA,uk;q=0.9,"
        "ru;q=0.8,en;q=0.7"
    )
}


session = requests.Session()

session.headers.update(
    HEADERS
)


# ==========================================================
# LOCK
# ==========================================================

def is_locked():

    if not os.path.exists(
        LOCK_FILE
    ):
        return False

    try:

        age = (
            time.time()
            - os.path.getmtime(
                LOCK_FILE
            )
        )

        if age > 3600:

            os.remove(
                LOCK_FILE
            )

            return False

        return True

    except Exception:

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

            try:

                os.remove(
                    LOCK_FILE
                )

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


            response = session.get(
                FEED_URL,
                timeout=60
            )


            if response.status_code != 200:

                last_error = (
                    f"HTTP "
                    f"{response.status_code}"
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


            return content

        except Exception as e:

            last_error = str(e)



            time.sleep(2)

    raise RuntimeError(
        "Не удалось загрузить XML: "
        + last_error
    )


# ==========================================================
# XML HELPERS
# ==========================================================

def local_name(tag):

    if not isinstance(
        tag,
        str
    ):
        return ""

    if "}" in tag:

        return tag.split(
            "}",
            1
        )[1]

    return tag


def get_child_text(
    element,
    wanted_name
):

    wanted_name = (
        wanted_name.lower()
    )

    for child in element:

        if (
            local_name(
                child.tag
            ).lower()
            == wanted_name
        ):

            return clean(
                child.text
            )

    return ""


def get_params(offer):

    result = {}

    for child in offer:

        if (
            local_name(
                child.tag
            ).lower()
            != "param"
        ):
            continue

        name = clean(
            child.attrib.get(
                "name",
                ""
            )
        )

        value = clean(
            child.text
        )

        if name:

            result[
                name.lower()
            ] = value

    return result


# ==========================================================
# PARSE XML
# ==========================================================

def parse_xml(xml_content):

    try:

        root = ET.fromstring(
            xml_content
        )

    except ET.ParseError as e:

        raise RuntimeError(
            f"XML повреждён: {e}"
        )

    offers = [
        element
        for element in root.iter()
        if local_name(
            element.tag
        ).lower() == "offer"
    ]

    if not offers:

        raise RuntimeError(
            "В XML отсутствуют <offer>"
        )


    if PRODUCT_LIMIT:

        offers = offers[
            :PRODUCT_LIMIT
        ]


    result = []

    for offer in offers:

        vendor_code = get_child_text(
            offer,
            "vendorCode"
        )

        params = get_params(
            offer
        )

        product_code = params.get(
            "код",
            ""
        )

        title = get_child_text(
            offer,
            "name"
        )

        url = get_child_text(
            offer,
            "url"
        )

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

        result.append({
            "CODE": vendor_code,
            "PRODUCT_CODE": product_code,
            "TITLE": title,
            "PRICE": "",
            "STATUS": status,
            "URL": url
        })

    if not result:

        raise RuntimeError(
            "После обработки XML "
            "не получено товаров"
        )

    return result


# ==========================================================
# LOGIN
# ==========================================================

def login_kmt():

    print("🔐 Авторизация...")

    try:

        response = session.get(
            LOGIN_URL,
            timeout=30
        )

        if response.status_code != 200:

            raise RuntimeError(
                f"Страница входа: HTTP {response.status_code}"
            )

        login_headers = {
            "Referer": LOGIN_URL,
            "Origin": "https://kmt5.com.ua",
            "X-Requested-With": "XMLHttpRequest",
            "Content-Type": (
                "application/x-www-form-urlencoded; "
                "charset=UTF-8"
            ),
            "Accept": (
                "application/json,"
                "text/javascript,"
                "text/html,"
                "*/*;q=0.01"
            )
        }

        payload = {
            "email": KMT_LOGIN,
            "password": KMT_PASSWORD
        }

        response = session.post(
            LOGIN_AJAX_URL,
            data=payload,
            headers=login_headers,
            timeout=30
        )

        if response.status_code != 200:

            raise RuntimeError(
                f"Авторизация: HTTP {response.status_code}"
            )

        try:

            data = response.json()

        except Exception:

            data = {}

        if data.get("success") is not True:

            raise RuntimeError(
                "КМТ не подтвердил авторизацию"
            )

        check = session.get(
            "https://kmt5.com.ua/my-account/",
            timeout=30
        )

        if check.status_code != 200:

            raise RuntimeError(
                "Не удалось проверить авторизацию"
            )

        print("✅ Авторизация успешна")

        return True

    except Exception as e:

        raise RuntimeError(
            f"Ошибка авторизации: {e}"
        )


# ==========================================================
# PRICE
# ==========================================================

def get_site_price(url):

    if not url:
        return ""

    try:

        response = session.get(
            url,
            timeout=15
        )

        if response.status_code != 200:
            return ""

        if len(response.text.strip()) < 500:
            return ""

        soup = BeautifulSoup(
            response.text,
            "html.parser"
        )

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

    except Exception:

        return ""

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

        ws.append([
            item["CODE"],
            item["PRODUCT_CODE"],
            item["TITLE"],
            item["PRICE"],
            item["STATUS"],
            item["URL"]
        ])

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



# ==========================================================
# RUN
# ==========================================================

def run_parser():

    if is_locked():


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

        # ==================================================
        # XML
        # ==================================================

        try:

            xml_content = download_feed()

            products = parse_xml(
                xml_content
            )

        except Exception as e:

            error_message = str(e)

            print(
                f"❌ ОШИБКА XML: "
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
        # LOGIN
        # ==================================================

        try:

            login_kmt()

        except Exception as e:

            error_message = str(e)

            print(
                f"❌ ОШИБКА АВТОРИЗАЦИИ: "
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
        # PRODUCTS
        # ==================================================

        total = len(products)

        price_found = 0

        price_missing = 0


        for index, product in enumerate(
            products,
            1
        ):

            title = product[
                "TITLE"
            ]



            price = get_site_price(
                product["URL"]
            )


            if price:

                price_found += 1



            else:

                price_missing += 1

  


            product[
                "PRICE"
            ] = price


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


        # ==================================================
        # SAVE
        # ==================================================

        if not products:

            error_message = (
                "Парсер завершился: "
                "товаров получено 0"
            )

            print(
                f"❌ {error_message}"
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


        try:

            save_excel(
                products
            )

        except Exception as e:

            error_message = (
                f"Ошибка сохранения Excel: {e}"
            )

            print(
                f"❌ {error_message}"
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



    finally:

        set_lock(False)


# ==========================================================
# MAIN
# ==========================================================

if __name__ == "__main__":

    run_parser()
