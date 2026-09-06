import os
import sys
import json
import time
import requests

from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

from bs4 import BeautifulSoup
from openpyxl import Workbook


USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Maraton")


# =========================================================
# SETTINGS
# =========================================================

BASE_URL = "https://maraton.ua"

LOGIN_URL = f"{BASE_URL}/login/"
START_URL = f"{BASE_URL}/ua/"

# ТЕСТ
PAGE_LIMIT = 2

# ПОЛНЫЙ ЗАПУСК
# PAGE_LIMIT = None


# =========================================================
# LOGIN
# =========================================================

# Задаём снаружи:
#
# Windows:
# set MARATON_LOGIN=...
# set MARATON_PASSWORD=...
#
# Linux:
# export MARATON_LOGIN=...
# export MARATON_PASSWORD=...

MARATON_LOGIN = os.getenv("MARATON_LOGIN", "").strip()
MARATON_PASSWORD = os.getenv("MARATON_PASSWORD", "").strip()


# =========================================================
# PATHS
# =========================================================

OUTPUT_DIR = os.path.abspath(
    "output/Maraton"
)

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
# SESSION
# =========================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/140.0.0.0 Safari/537.36"
    ),
    "Accept-Language": (
        "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7"
    ),
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

        age = (
            time.time()
            - os.path.getmtime(LOCK_FILE)
        )

        if age > 3600:

            os.remove(LOCK_FILE)

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


# =========================================================
# REQUEST
# =========================================================

def get_page(url):

    for attempt in range(3):

        try:

            response = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            response.raise_for_status()

            return response.text

        except Exception:

            if attempt < 2:
                time.sleep(2)

    return None


# =========================================================
# LOGIN
# =========================================================

def login():

    if not MARATON_LOGIN:
        raise Exception(
            "Не задан MARATON_LOGIN"
        )

    if not MARATON_PASSWORD:
        raise Exception(
            "Не задан MARATON_PASSWORD"
        )

    print("🔐 Авторизация...")

    # -----------------------------------------------------
    # Сначала открываем страницу логина.
    # Это нужно для cookies / сессии.
    # -----------------------------------------------------

    try:

        response = session.get(
            LOGIN_URL,
            timeout=30
        )

        response.raise_for_status()

    except Exception as e:

        raise Exception(
            f"Не удалось открыть страницу входа: {e}"
        )

    # -----------------------------------------------------
    # POST LOGIN
    # -----------------------------------------------------

    payload = {
        "login": MARATON_LOGIN,
        "password": MARATON_PASSWORD,
        "remember": "1",
        "wa_auth_login": "1",
    }

    headers = {
        "Referer": LOGIN_URL,
        "X-Requested-With": "XMLHttpRequest",
        "Accept": "application/json, text/javascript, */*; q=0.01",
    }

    try:

        response = session.post(
            LOGIN_URL,
            data=payload,
            headers=headers,
            timeout=30,
            allow_redirects=True
        )

    except Exception as e:

        raise Exception(
            f"Ошибка авторизации: {e}"
        )

    # -----------------------------------------------------
    # ПРОВЕРЯЕМ СЕССИЮ
    # -----------------------------------------------------

    check = session.get(
        START_URL,
        timeout=30
    )

    check.raise_for_status()

    soup = BeautifulSoup(
        check.text,
        "html.parser"
    )

    page_text = soup.get_text(
        " ",
        strip=True
    )

    # Если на странице всё ещё есть "Вход",
    # проверяем дополнительные признаки
    # авторизации.

    logged_in = False

    # После авторизации обычно появляется
    # личный кабинет / выход.
    login_links = soup.select(
        'a[href="/login/"]'
    )

    logout_links = soup.select(
        'a[href*="logout"]'
    )

    personal_links = soup.select(
        'a[href*="/my/"]'
    )

    if logout_links or personal_links:
        logged_in = True

    # -----------------------------------------------------
    # Дополнительная проверка по цене
    # -----------------------------------------------------

    price = soup.select_one(
        ".price-number.s-product-price"
    )

    # Если цена уже появилась — авторизация точно прошла.
    if price:
        logged_in = True

    if not logged_in:

        # Возможный JSON ответ
        try:

            data = response.json()

            text = json.dumps(
                data,
                ensure_ascii=False
            ).lower()

            if (
                "success" in text
                and "error" not in text
            ):
                logged_in = True

        except Exception:
            pass

    if not logged_in:

        raise Exception(
            "Не удалось подтвердить авторизацию Maraton"
        )

    print("✅ Авторизация успешна")


# =========================================================
# URL
# =========================================================

def normalize_url(url):

    if not url:
        return ""

    url = url.strip()

    if url.startswith("/"):

        url = urljoin(
            BASE_URL,
            url
        )

    if url.startswith(
        "http://maraton.ua"
    ):

        url = url.replace(
            "http://maraton.ua",
            "https://maraton.ua",
            1
        )

    return url.split("#")[0]


# =========================================================
# CATEGORY LINKS
# =========================================================

def get_category_links(html):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    links = set()

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = normalize_url(
            a.get("href")
        )

        if not href:
            continue

        parsed = urlparse(href)

        if parsed.netloc not in (
            "maraton.ua",
            "www.maraton.ua"
        ):
            continue

        path = parsed.path.rstrip("/")

        if not path.startswith("/ua/"):
            continue

        # -------------------------------------------------
        # Исключаем служебные страницы
        # -------------------------------------------------

        excluded = (
            "/ua/search",
            "/ua/login",
            "/ua/signup",
            "/ua/forgotpassword",
            "/ua/cart",
            "/ua/checkout",
            "/ua/opt",
            "/ua/about",
            "/ua/contact",
        )

        if any(
            path.startswith(x)
            for x in excluded
        ):
            continue

        # -------------------------------------------------
        # Ссылка должна быть частью каталога
        # -------------------------------------------------

        if path.count("/") >= 2:

            links.add(href)

    return sorted(links)


# =========================================================
# PAGINATION URL
# =========================================================

def make_page_url(
    url,
    page
):

    parsed = urlparse(url)

    query = parse_qs(
        parsed.query
    )

    query["page"] = [str(page)]

    new_query = urlencode(
        query,
        doseq=True
    )

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        new_query,
        parsed.fragment
    ))


# =========================================================
# PRODUCT CARDS
# =========================================================

def parse_product_cards(html):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    products = []

    # -----------------------------------------------------
    # Ищем реальные ссылки на товары
    # -----------------------------------------------------

    candidates = []

    for a in soup.find_all(
        "a",
        href=True
    ):

        href = normalize_url(
            a.get("href")
        )

        if not href:
            continue

        path = urlparse(
            href
        ).path

        if not path.startswith("/"):
            continue

        if any(
            x in path
            for x in (
                "/login",
                "/signup",
                "/search",
                "/cart",
                "/checkout",
                "/forgotpassword",
            )
        ):
            continue

        candidates.append(
            (a, href)
        )

    # -----------------------------------------------------
    # Проверяем родительские блоки
    # -----------------------------------------------------

    seen_urls = set()

    for a, href in candidates:

        if href in seen_urls:
            continue

        parent = a

        for _ in range(6):

            if parent is None:
                break

            text = parent.get_text(
                " ",
                strip=True
            )

            # У товара должен присутствовать
            # либо цена, либо наличие.
            if (
                "$" in text
                or "Опт" in text
                or "В наличии" in text
                or "Немає в наявності" in text
            ):

                seen_urls.add(href)

                product = parse_product_card(
                    parent,
                    href
                )

                if product:

                    products.append(
                        product
                    )

                break

            parent = parent.parent

    return products


# =========================================================
# PRODUCT CARD
# =========================================================

def parse_product_card(
    block,
    url
):

    # -----------------------------------------------------
    # TITLE
    # -----------------------------------------------------

    title = ""

    title_element = block.select_one(
        "h3, h2, h4, .item-name, .product-name"
    )

    if title_element:

        title = title_element.get_text(
            " ",
            strip=True
        )

    if not title:

        # fallback — ищем ссылку
        # с максимально подходящим текстом

        links = block.find_all(
            "a"
        )

        for link in links:

            text = link.get_text(
                " ",
                strip=True
            )

            if (
                text
                and len(text) > 3
                and text != "Подробнее"
            ):

                title = text

                break

    if not title:
        return None

    # -----------------------------------------------------
    # PRICE
    # -----------------------------------------------------

    price = ""

    price_element = block.select_one(
        ".price-number.s-product-price"
    )

    if price_element:

        price = (
            price_element.get(
                "data-price"
            )
            or price_element.get_text(
                " ",
                strip=True
            )
        ).strip()

    if not price:

        price_element = block.select_one(
            ".item-price, .price"
        )

        if price_element:

            price = price_element.get_text(
                " ",
                strip=True
            )

    # -----------------------------------------------------
    # STOCK
    # -----------------------------------------------------

    status = ""

    stock = block.select_one(
        ".stock-text"
    )

    if stock:

        status = stock.get_text(
            " ",
            strip=True
        )

    if not status:

        text = block.get_text(
            " ",
            strip=True
        )

        if "В наличии" in text:

            status = "В наличии"

        elif "Немає в наявності" in text:

            status = "Нет в наличии"

    # -----------------------------------------------------
    # PRODUCT ID
    # -----------------------------------------------------

    product_id = ""

    product_id_element = block.select_one(
        'input[name="product_id"]'
    )

    if product_id_element:

        product_id = (
            product_id_element.get(
                "value"
            )
            or ""
        ).strip()

    # -----------------------------------------------------
    # IMAGE
    # -----------------------------------------------------

    image = ""

    image_element = block.select_one(
        "img"
    )

    if image_element:

        image = (
            image_element.get(
                "data-src"
            )
            or image_element.get(
                "src"
            )
            or ""
        )

        image = normalize_url(
            image
        )

    return {
        "article": product_id,
        "title": title,
        "price": price,
        "currency": "USD",
        "status": status,
        "url": url,
        "image": image,
    }


# =========================================================
# RUN
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
        # LOGIN
        # -------------------------------------------------

        login()

        # -------------------------------------------------
        # MAIN PAGE
        # -------------------------------------------------

        html = get_page(
            START_URL
        )

        if not html:

            raise Exception(
                "Не удалось открыть каталог Maraton"
            )

        # -------------------------------------------------
        # CATEGORIES
        # -------------------------------------------------

        category_links = get_category_links(
            html
        )

        print(
            f"📂 Categories: "
            f"{len(category_links)}"
        )

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
            "Ссылка",
            "Изображение"
        ])

        # -------------------------------------------------
        # DEDUP
        # -------------------------------------------------

        seen = set()

        # -------------------------------------------------
        # TEST
        # -------------------------------------------------

        categories = category_links

        if PAGE_LIMIT:

            categories = categories[
                :PAGE_LIMIT
            ]

            print(
                f"🧪 TEST MODE: "
                f"{len(categories)} categories"
            )

        # -------------------------------------------------
        # CATEGORIES
        # -------------------------------------------------

        for category_index, category_url in enumerate(
            categories,
            1
        ):

            print(
                f"🔹 {category_url}"
            )

            page = 1

            while True:

                page_url = make_page_url(
                    category_url,
                    page
                )

                html = get_page(
                    page_url
                )

                if not html:
                    break

                products = parse_product_cards(
                    html
                )

                if not products:
                    break

                new_products = 0

                for product in products:

                    article = product[
                        "article"
                    ]

                    url = product[
                        "url"
                    ]

                    title = product[
                        "title"
                    ]

                    key = (
                        article
                        or url
                        or title
                    )

                    if key in seen:
                        continue

                    seen.add(key)

                    ws.append([
                        product["article"],
                        product["title"],
                        product["price"],
                        product["currency"],
                        product["status"],
                        product["url"],
                        product["image"]
                    ])

                    new_products += 1

                # -------------------------------------------------
                # Если страница ничего нового не дала —
                # скорее всего пагинация закончилась.
                # -------------------------------------------------

                if new_products == 0:
                    break

                page += 1

                # -------------------------------------------------
                # Защита
                # -------------------------------------------------

                if page > 200:
                    break

        # -------------------------------------------------
        # SAVE
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

        save_status(
            False,
            100,
            USER,
            FILE_PATH
        )

        print(
            f"📦 Products: {len(seen)}"
        )

        print(
            "✅ Готово. Maraton"
        )

    finally:

        set_lock(False)


# =========================================================
# START
# =========================================================

if __name__ == "__main__":
    run_parser()
