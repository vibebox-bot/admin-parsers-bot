
import json
import time
import re
from pathlib import Path

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ==========================================================
# НАСТРОЙКИ
# ==========================================================

BASE_URL = "https://b2b.delltalife.com"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"

CATEGORY_LIMIT = 2

OUTPUT_DIR = Path("output/Dellta")
OUTPUT_FILE = OUTPUT_DIR / "Dellta_LIVE.xlsx"
STATUS_FILE = OUTPUT_DIR / "status.json"
LOCK_FILE = OUTPUT_DIR / "lock.txt"

SHEET_NAME = "Dellta"


# ==========================================================
# LOG
# ==========================================================

def log(message):
    print(message, flush=True)


# ==========================================================
# СЛУЖЕБНОЕ
# ==========================================================

def clean_text(value):
    if not value:
        return ""

    return re.sub(r"\s+", " ", str(value)).strip()


def clean_price(value):
    if not value:
        return ""

    value = value.replace("$", "")
    value = value.replace("USD", "")
    value = value.replace(",", ".")
    value = clean_text(value)

    match = re.search(r"\d+(?:\.\d+)?", value)

    if not match:
        return ""

    return float(match.group())


def normalize_url(url):
    if not url:
        return ""

    if url.startswith("http://") or url.startswith("https://"):
        return url

    if url.startswith("/"):
        return BASE_URL + url

    return BASE_URL + "/" + url


# ==========================================================
# STATUS
# ==========================================================

def save_status(status, **extra):
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    data = {
        "status": status,
        **extra
    }

    STATUS_FILE.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )


def create_lock():
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    LOCK_FILE.write_text(
        "running",
        encoding="utf-8"
    )


def remove_lock():
    try:
        LOCK_FILE.unlink()
    except FileNotFoundError:
        pass


# ==========================================================
# EXCEL
# ==========================================================

def create_excel():
    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    wb = Workbook()

    ws = wb.active
    ws.title = SHEET_NAME

    ws.append([
        "SKU",
        "TITLE",
        "PRICE",
        "STATUS",
        "URL"
    ])

    wb.save(OUTPUT_FILE)


def reset_excel():
    """
    Полностью очищает Excel и создаёт новый
    только с заголовками.
    """

    try:
        if OUTPUT_FILE.exists():
            OUTPUT_FILE.unlink()

        create_excel()

    except Exception:
        pass


def append_products(products):

    if not products:
        return

    if not OUTPUT_FILE.exists():
        create_excel()

    wb = load_workbook(
        OUTPUT_FILE
    )

    ws = wb[SHEET_NAME]

    for product in products:

        ws.append([
            product["SKU"],
            product["TITLE"],
            product["PRICE"],
            product["STATUS"],
            product["URL"]
        ])

    wb.save(OUTPUT_FILE)


# ==========================================================
# ПАРСИНГ ТОВАРОВ
# ==========================================================

def parse_products(html):

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    rows = soup.select(
        "tr.itemPosition.simple"
    )

    products = []

    for row in rows:

        # --------------------------------------------------
        # SKU
        # --------------------------------------------------

        sku_el = row.select_one(
            "td.td_2 .gray"
        )

        sku = clean_text(
            sku_el.get_text(
                " ",
                strip=True
            )
            if sku_el
            else ""
        )

        # --------------------------------------------------
        # TITLE + URL
        # --------------------------------------------------

        title_el = row.select_one(
            "td.td_2 a[href*='/invertoryi-']"
        )

        if not title_el:

            title_el = row.select_one(
                "td.td_2 a[href]"
            )

        title = clean_text(
            title_el.get_text(
                " ",
                strip=True
            )
            if title_el
            else ""
        )

        url = normalize_url(
            title_el.get("href", "")
            if title_el
            else ""
        )

        # --------------------------------------------------
        # STATUS
        # --------------------------------------------------

        status_el = row.select_one(
            "td.td_2 .are-available, "
            "td.td_2 .not-available, "
            "td.td_2 div[class^='are-']"
        )

        status = clean_text(
            status_el.get_text(
                " ",
                strip=True
            )
            if status_el
            else ""
        )

        # --------------------------------------------------
        # ЦЕНА КОМП. ДИЛЕР
        # --------------------------------------------------

        price_el = row.select_one(
            "td.td_3 .price-table "
            "tr.line-1 span.active"
        )

        if not price_el:

            price_el = row.select_one(
                "td.td_3 tr.line-1 span.active"
            )

        price = clean_price(
            price_el.get_text(
                " ",
                strip=True
            )
            if price_el
            else ""
        )

        # --------------------------------------------------
        # ПРОПУСК ПУСТЫХ СТРОК
        # --------------------------------------------------

        if not sku and not title:
            continue

        products.append({
            "SKU": sku,
            "TITLE": title,
            "PRICE": price,
            "STATUS": status,
            "URL": url
        })

    return products


# ==========================================================
# КАТЕГОРИИ
# ==========================================================

def get_categories(page):

    page.goto(
        BASE_URL + "/",
        wait_until="domcontentloaded",
        timeout=60000
    )

    time.sleep(5)

    html = page.content()

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    categories = []

    for link in soup.select(
        "a[href*='/invertoryi-']"
    ):

        href = link.get(
            "href",
            ""
        )

        name = clean_text(
            link.get_text(
                " ",
                strip=True
            )
        )

        if not href or not name:
            continue

        url = normalize_url(href)

        item = {
            "name": name,
            "url": url
        }

        if item not in categories:
            categories.append(item)

    return categories


# ==========================================================
# ОТКРЫТИЕ КАТЕГОРИИ
# ==========================================================

def open_category(page, category):

    page.goto(
        category["url"],
        wait_until="domcontentloaded",
        timeout=60000
    )

    time.sleep(3)

    return page.content()


# ==========================================================
# АВТОРИЗАЦИЯ
# ==========================================================

def login(page):

    page.goto(
        BASE_URL + "/",
        wait_until="domcontentloaded",
        timeout=60000
    )

    time.sleep(5)

    enter_button = page.locator(
        "#a-enter"
    )

    if not enter_button.count():
        return False

    enter_button.click()

    try:

        page.locator(
            "#login-form"
        ).wait_for(
            state="visible",
            timeout=15000
        )

    except PlaywrightTimeoutError:

        return False

    page.locator(
        "input[name='email_auth']"
    ).fill(
        EMAIL
    )

    page.locator(
        "input[name='pass_auth']"
    ).fill(
        PASSWORD
    )

    login_button = page.locator(
        "#login-form button.modalButton"
    )

    if not login_button.count():
        return False

    login_button.click()

    time.sleep(5)

    try:

        if page.locator(
            "#login-form"
        ).is_visible():

            return False

    except Exception:

        pass

    return True


# ==========================================================
# ОСНОВНОЙ ПАРСЕР
# ==========================================================

def run_parser():

    create_lock()

    save_status(
        "running",
        products=0,
        categories=0
    )

    log("🚀 Запуск парсера Dellta")

    try:

        with sync_playwright() as p:

            # ------------------------------------------------
            # FIREFOX
            # ------------------------------------------------

            browser = p.firefox.launch(
                headless=True
            )

            context = browser.new_context(
                viewport={
                    "width": 1440,
                    "height": 900
                },
                locale="uk-UA",
                timezone_id="Europe/Kyiv",
                user_agent=(
                    "Mozilla/5.0 "
                    "(X11; Linux x86_64; rv:128.0) "
                    "Gecko/20100101 Firefox/128.0"
                )
            )

            page = context.new_page()

            # ------------------------------------------------
            # EXCEL
            # ------------------------------------------------

            reset_excel()

            # ------------------------------------------------
            # КАТЕГОРИИ
            # ------------------------------------------------

            categories = get_categories(
                page
            )

            log(
                f"📂 Категорий: {len(categories)}"
            )

            if not categories:

                reset_excel()

                save_status(
                    "error",
                    products=0,
                    categories=0,
                    error="Categories not found"
                )

                browser.close()

                return

            if CATEGORY_LIMIT:
                categories = categories[
                    :CATEGORY_LIMIT
                ]

            # ------------------------------------------------
            # АВТОРИЗАЦИЯ
            # ------------------------------------------------

            if not login(page):

                reset_excel()

                save_status(
                    "error",
                    products=0,
                    categories=0,
                    error="Authorization failed"
                )

                browser.close()

                return

            # ------------------------------------------------
            # ПАРСИНГ
            # ------------------------------------------------

            total_products = 0
            successful_categories = 0

            for index, category in enumerate(
                categories,
                start=1
            ):

                try:

                    html = open_category(
                        page,
                        category
                    )

                    products = parse_products(
                        html
                    )

                    if products:

                        append_products(
                            products
                        )

                        total_products += len(
                            products
                        )

                        successful_categories += 1

                        log(
                            f"📦 Товаров: {len(products)}"
                        )

                    save_status(
                        "running",
                        products=total_products,
                        categories=index
                    )

                except Exception:

                    continue

            # ------------------------------------------------
            # ЕСЛИ НИ ОДНА КАТЕГОРИЯ НЕ ОБРАБОТАЛАСЬ
            # ------------------------------------------------

            if successful_categories == 0:

                reset_excel()

                save_status(
                    "error",
                    products=0,
                    categories=0,
                    error="All categories failed"
                )

                browser.close()

                return

            # ------------------------------------------------
            # ЗАВЕРШЕНИЕ
            # ------------------------------------------------

            browser.close()

            save_status(
                "completed",
                products=total_products,
                categories=successful_categories,
                file=str(OUTPUT_FILE)
            )

            log(
                "🎉 DELLTA ЗАВЕРШЕН"
            )

    except Exception as e:

        # ----------------------------------------------------
        # ЛЮБАЯ КРИТИЧЕСКАЯ ОШИБКА
        # ----------------------------------------------------

        reset_excel()

        log(
            f"❌ Критическая ошибка: {e}"
        )

        save_status(
            "error",
            products=0,
            categories=0,
            error=str(e)
        )

    finally:

        remove_lock()


# ==========================================================
# START
# ==========================================================

if __name__ == "__main__":
    run_parser()
