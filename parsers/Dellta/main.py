import os
import sys
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

EMAIL = os.getenv("DELLTA_EMAIL", "angelinatitor@gmail.com")
PASSWORD = os.getenv("DELLTA_PASSWORD", "")

CATEGORY_LIMIT = 2

OUTPUT_DIR = Path("output/Dellta")
OUTPUT_FILE = OUTPUT_DIR / "Dellta_LIVE.xlsx"
STATUS_FILE = OUTPUT_DIR / "status.json"
LOCK_FILE = OUTPUT_DIR / "lock.txt"


# ==========================================================
# СЛУЖЕБНЫЕ ФУНКЦИИ
# ==========================================================

def log(message):
    print(message, flush=True)


def save_status(status, **extra):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    data = {
        "status": status,
        "time": time.strftime("%Y-%m-%d %H:%M:%S"),
    }

    data.update(extra)

    try:
        with open(STATUS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def create_lock():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    try:
        LOCK_FILE.write_text(
            str(os.getpid()),
            encoding="utf-8"
        )
    except Exception:
        pass


def remove_lock():
    try:
        if LOCK_FILE.exists():
            LOCK_FILE.unlink()
    except Exception:
        pass


def clean_text(value):
    if value is None:
        return ""

    return re.sub(
        r"\s+",
        " ",
        str(value)
    ).strip()


def clean_price(value):
    if not value:
        return ""

    value = value.replace("$", "")
    value = value.replace("USD", "")
    value = value.replace(",", ".")
    value = clean_text(value)

    match = re.search(
        r"\d+(?:\.\d+)?",
        value
    )

    if not match:
        return ""

    try:
        return float(match.group(0))
    except Exception:
        return ""


def normalize_url(url):
    if not url:
        return ""

    if url.startswith("http://") or url.startswith("https://"):
        return url

    if not url.startswith("/"):
        url = "/" + url

    return BASE_URL + url


# ==========================================================
# EXCEL
# ==========================================================

def create_excel():
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Dellta"

    ws.append([
        "SKU",
        "TITLE",
        "PRICE",
        "STATUS",
        "URL"
    ])

    wb.save(OUTPUT_FILE)


def append_products(products):
    if not products:
        return

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if OUTPUT_FILE.exists():
        wb = load_workbook(OUTPUT_FILE)
        ws = wb["Dellta"]
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = "Dellta"

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

    existing_sku = set()

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):
        sku = row[0]

        if sku:
            existing_sku.add(
                str(sku).strip()
            )

    added = 0

    for product in products:
        sku = str(
            product.get("SKU", "")
        ).strip()

        if not sku:
            continue

        if sku in existing_sku:
            continue

        ws.append([
            product.get("SKU", ""),
            product.get("TITLE", ""),
            product.get("PRICE", ""),
            product.get("STATUS", ""),
            product.get("URL", ""),
        ])

        existing_sku.add(sku)
        added += 1

    wb.save(OUTPUT_FILE)

    log(
        f"💾 Добавлено в Excel: {added}"
    )


# ==========================================================
# ПАРСИНГ ТОВАРОВ
# ==========================================================

def parse_products(html):
    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    products = []

    rows = soup.select(
        "tr.itemPosition.simple"
    )

    log(
        f"📦 Найдено строк товаров: {len(rows)}"
    )

    for row in rows:

        # --------------------------------------------------
        # SKU
        # --------------------------------------------------

        sku_el = row.select_one(
            "td.td_2 .gray"
        )

        sku = clean_text(
            sku_el.get_text(" ", strip=True)
            if sku_el
            else ""
        )

        # --------------------------------------------------
        # TITLE + URL
        # --------------------------------------------------

        title_el = row.select_one(
            "td.td_2 a[href]"
        )

        if not title_el:
            title_el = row.select_one(
                'td.td_2 a[href*="/invertoryi-"]'
            )

        if not title_el:
            continue

        title = clean_text(
            title_el.get_text(
                " ",
                strip=True
            )
        )

        url = normalize_url(
            title_el.get("href", "")
        )

        # --------------------------------------------------
        # STATUS
        # --------------------------------------------------

        status_el = row.select_one(
            "td.td_2 .are-available"
        )

        if not status_el:
            status_el = row.select_one(
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
        # ЦЕНА ДИЛЕР
        #
        # Комп. ДИЛЕР = line-1
        # Берём только span.active
        #
        # Например:
        #
        # <tr class="line-1">
        #   <td>Комп. ДИЛЕР</td>
        #   <td>
        #       <span class="active">12.80 $</span>
        #       <span>576.00 ₴</span>
        #   </td>
        # </tr>
        # --------------------------------------------------

        price_el = row.select_one(
            "td.td_3 .price-table tr.line-1 span.active"
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

        product = {
            "SKU": sku,
            "TITLE": title,
            "PRICE": price,
            "STATUS": status,
            "URL": url,
        }

        products.append(product)

    return products


# ==========================================================
# ПОЛУЧЕНИЕ КАТЕГОРИЙ
# ==========================================================

def get_categories(page):
    log("📂 Получаем категории...")

    html = page.content()

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    categories = []

    # Основной вариант
    for a in soup.select(
        'a[href*="/invertoryi-"]'
    ):
        href = a.get("href", "")
        title = clean_text(
            a.get_text(
                " ",
                strip=True
            )
        )

        if not href or not title:
            continue

        url = normalize_url(href)

        item = {
            "title": title,
            "url": url,
        }

        if item not in categories:
            categories.append(item)

    # Убираем ссылки, которые являются товарами
    # если они явно выглядят как товарные карточки.
    #
    # Оставляем уникальные ссылки.

    unique = []
    seen = set()

    for category in categories:

        url = category["url"]

        if url in seen:
            continue

        seen.add(url)
        unique.append(category)

    categories = unique

    log(
        f"📂 Найдено категорий/ссылок: {len(categories)}"
    )

    if CATEGORY_LIMIT:
        categories = categories[:CATEGORY_LIMIT]

        log(
            f"🧪 Тестовый лимит категорий: "
            f"{CATEGORY_LIMIT}"
        )

    return categories


# ==========================================================
# ОТКРЫТИЕ КАТЕГОРИИ
# ==========================================================

def open_category(page, category):
    title = category["title"]
    url = category["url"]

    log("")
    log(
        f"📂 Категория: {title}"
    )

    log(
        f"🌐 URL: {url}"
    )

    try:
        response = page.goto(
            url,
            wait_until="domcontentloaded",
            timeout=60000
        )

        status = (
            response.status
            if response
            else "UNKNOWN"
        )

        log(
            f"🌐 HTTP: {status}"
        )

        page.wait_for_timeout(2000)

        return True

    except Exception as e:
        log(
            f"❌ Ошибка открытия категории: {e}"
        )

        return False


# ==========================================================
# АВТОРИЗАЦИЯ
# ==========================================================

def login(page):
    log("🔐 Авторизация Dellta...")

    log(
        f"🌐 GET {BASE_URL}/"
    )

    try:
        response = page.goto(
            BASE_URL + "/",
            wait_until="domcontentloaded",
            timeout=60000
        )

        http_status = (
            response.status
            if response
            else "UNKNOWN"
        )

        log(
            f"🌐 HTTP: {http_status}"
        )

        log(
            f"🌐 URL: {page.url}"
        )

    except Exception as e:
        log(
            f"⚠ Ошибка GET: {e}"
        )

    log(
        "⏳ Ждём выполнение JavaScript..."
    )

    page.wait_for_timeout(5000)

    # ------------------------------------------------------
    # ПРОВЕРЯЕМ КНОПКУ ВХОДА
    # ------------------------------------------------------

    login_link = page.locator(
        "#a-enter"
    )

    try:
        login_link.wait_for(
            state="visible",
            timeout=30000
        )

        log(
            "✅ Кнопка «Вхід/Реєстрація» найдена"
        )

    except PlaywrightTimeoutError:
        log(
            "❌ Кнопка #a-enter не появилась"
        )

        # Диагностика
        try:
            log(
                f"📄 Текущий title: {page.title()}"
            )

            log(
                f"🌐 Текущий URL: {page.url}"
            )
        except Exception:
            pass

        return False

    # ------------------------------------------------------
    # ОТКРЫВАЕМ MODAL
    # ------------------------------------------------------

    log(
        "🔑 Открываем «Вхід/Реєстрація»..."
    )

    try:
        login_link.click(
            timeout=30000
        )

        log(
            "✅ Кнопка нажата"
        )

    except Exception as e:
        log(
            f"❌ Не удалось нажать кнопку входа: {e}"
        )

        return False

    # ------------------------------------------------------
    # ЖДЁМ ФОРМУ
    # ------------------------------------------------------

    login_form = page.locator(
        "#login-form"
    )

    try:
        login_form.wait_for(
            state="visible",
            timeout=30000
        )

        log(
            "✅ Форма авторизации открыта"
        )

    except PlaywrightTimeoutError:

        log(
            "❌ #login-form не стала видимой"
        )

        # Проверяем существование
        try:
            count = login_form.count()

            log(
                f"🔎 #login-form элементов в DOM: {count}"
            )
        except Exception:
            pass

        return False

    # ------------------------------------------------------
    # EMAIL
    # ------------------------------------------------------

    email_input = login_form.locator(
        'input[name="email_auth"]'
    )

    log(
        "✏ Заполняем email..."
    )

    try:
        email_input.fill(
            EMAIL,
            timeout=30000
        )

    except Exception as e:
        log(
            f"❌ Ошибка email: {e}"
        )

        return False

    # ------------------------------------------------------
    # PASSWORD
    # ------------------------------------------------------

    password_input = login_form.locator(
        'input[name="pass_auth"]'
    )

    log(
        "✏ Заполняем пароль..."
    )

    try:
        password_input.fill(
            PASSWORD,
            timeout=30000
        )

    except Exception as e:
        log(
            f"❌ Ошибка password: {e}"
        )

        return False

    # ------------------------------------------------------
    # SUBMIT
    # ------------------------------------------------------

    login_button = login_form.locator(
        "button.modalButton"
    )

    log(
        "🔐 Нажимаем «Увійти»..."
    )

    try:
        login_button.click(
            timeout=30000
        )

    except Exception as e:
        log(
            f"❌ Ошибка кнопки входа: {e}"
        )

        return False

    # ------------------------------------------------------
    # ЖДЁМ AJAX
    # ------------------------------------------------------

    log(
        "⏳ Ждём завершения авторизации..."
    )

    page.wait_for_timeout(5000)

    # ------------------------------------------------------
    # ПРОВЕРКА
    # ------------------------------------------------------

    try:
        form_visible = login_form.is_visible()

    except Exception:
        form_visible = False

    if form_visible:
        log(
            "⚠ Форма входа всё ещё открыта"
        )

        # Иногда сервер отвечает ошибкой внутри modal.
        try:
            modal_text = clean_text(
                page.locator(
                    "#login-form"
                ).locator(
                    ".."
                ).inner_text()
            )

            if modal_text:
                log(
                    f"⚠ Ответ формы: {modal_text[:500]}"
                )
        except Exception:
            pass

        return False

    log(
        "✅ Авторизация завершена"
    )

    return True


# ==========================================================
# ОСНОВНОЙ ПАРСЕР
# ==========================================================

def run_parser():
    create_lock()

    save_status(
        "running"
    )

    log(
        "🔥 DELLTA LIFE PARSER"
    )

    log(
        "🚀 Запуск парсера Dellta"
    )

    if not PASSWORD:
        log(
            "⚠ ВНИМАНИЕ: DELLTA_PASSWORD не задан"
        )

    create_excel()

    try:

        with sync_playwright() as p:

            log(
                "🦊 Firefox Playwright найден"
            )

            log(
                "🦊 Запускаем Firefox..."
            )

            browser = p.firefox.launch(
                headless=True
            )

            context = browser.new_context(
                viewport={
                    "width": 1440,
                    "height": 900,
                },
                locale="uk-UA",
                timezone_id="Europe/Kyiv",
                user_agent=(
                    "Mozilla/5.0 "
                    "(X11; Linux x86_64; rv:128.0) "
                    "Gecko/20100101 Firefox/128.0"
                ),
            )

            page = context.new_page()

            # --------------------------------------------------
            # LOGIN
            # --------------------------------------------------

            if not login(page):

                save_status(
                    "error",
                    message="Авторизация Dellta не выполнена"
                )

                browser.close()

                return

            # --------------------------------------------------
            # КАТЕГОРИИ
            # --------------------------------------------------

            categories = get_categories(
                page
            )

            if not categories:

                log(
                    "❌ Категории не найдены"
                )

                save_status(
                    "error",
                    message="Категории не найдены"
                )

                browser.close()

                return

            total_products = 0

            # --------------------------------------------------
            # ПАРСИМ КАТЕГОРИИ
            # --------------------------------------------------

            for index, category in enumerate(
                categories,
                start=1
            ):

                log("")
                log(
                    f"📂 КАТЕГОРИЯ "
                    f"{index}/{len(categories)}"
                )

                if not open_category(
                    page,
                    category
                ):
                    continue

                html = page.content()

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

                    log(
                        f"✅ Товаров: "
                        f"{len(products)}"
                    )

                else:
                    log(
                        "⚠ Товары не найдены"
                    )

            # --------------------------------------------------
            # ЗАВЕРШЕНИЕ
            # --------------------------------------------------

            log("")
            log(
                "================================"
            )

            log(
                "🎉 DELLTA ЗАВЕРШЕН"
            )

            log(
                f"📦 Всего товаров: "
                f"{total_products}"
            )

            log(
                f"📄 Excel: "
                f"{OUTPUT_FILE}"
            )

            log(
                "================================"
            )

            save_status(
                "done",
                products=total_products
            )

            browser.close()

    except Exception as e:

        log(
            f"❌ Критическая ошибка: {e}"
        )

        save_status(
            "error",
            message=str(e)
        )

    finally:

        remove_lock()


# ==========================================================
# START
# ==========================================================

if __name__ == "__main__":
    run_parser()
