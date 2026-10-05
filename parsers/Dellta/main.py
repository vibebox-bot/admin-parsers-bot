import os
import json
import time
import re
from pathlib import Path
from datetime import datetime

from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ==========================================================
# НАСТРОЙКИ
# ==========================================================

BASE_URL = "https://b2b.delltalife.com"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"

# Максимальное количество категорий
# None = все категории
#CATEGORY_LIMIT = 1
CATEGORY_LIMIT = 2

OUTPUT_DIR = Path("output/Dellta")
OUTPUT_FILE = OUTPUT_DIR / "Dellta_LIVE.xlsx"
STATUS_FILE = OUTPUT_DIR / "status.json"
LOCK_FILE = OUTPUT_DIR / "lock.txt"

SHEET_NAME = "Dellta"

# USER приходит из run.py / bot.py
import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"


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

def save_status(
    running=False,
    progress=0,
    user="",
    file_path="",
    verification_required=False
):

    OUTPUT_DIR.mkdir(
        parents=True,
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
        "verification_required": verification_required
    }

    STATUS_FILE.write_text(
        json.dumps(
            data,
            ensure_ascii=False,
            indent=2
        ),
        encoding="utf-8"
    )


# ==========================================================
# LOCK
# ==========================================================

def is_locked():

    if not LOCK_FILE.exists():
        return False

    try:

        age = (
            time.time()
            - LOCK_FILE.stat().st_mtime
        )

        # Если lock старше часа — считаем его зависшим
        if age > 3600:

            LOCK_FILE.unlink(
                missing_ok=True
            )

            return False

        return True

    except Exception:

        return False


def set_lock(state):

    OUTPUT_DIR.mkdir(
        parents=True,
        exist_ok=True
    )

    if state:

        LOCK_FILE.write_text(
            str(time.time()),
            encoding="utf-8"
        )

    else:

        LOCK_FILE.unlink(
            missing_ok=True
        )


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

    wb.save(
        OUTPUT_FILE
    )


def reset_excel():

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

    wb.save(
        OUTPUT_FILE
    )


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
            title_el.get(
                "href",
                ""
            )
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
# КАТЕГОРИИ — ТОЛЬКО 1-Й УРОВЕНЬ
# ==========================================================

def get_categories(page):

    categories = []
    seen = set()

    try:

        # --------------------------------------------------
        # Открываем меню "Каталог товарів"
        # --------------------------------------------------

        catalog_button = page.locator(
            "a.catalogButton"
        ).first

        if catalog_button.count():

            try:
                catalog_button.click()
                time.sleep(1)
            except Exception:
                pass

        # --------------------------------------------------
        # Ждём существование меню
        # visible НЕ требуем
        # --------------------------------------------------

        page.wait_for_selector(
            "ul.firstUl",
            state="attached",
            timeout=15000
        )

        # --------------------------------------------------
        # Только ПРЯМЫЕ li первого уровня
        # --------------------------------------------------

        items = page.locator(
            "ul.firstUl > li"
        )

        count = items.count()

        for i in range(count):

            li = items.nth(i)

            classes = (
                li.get_attribute("class")
                or ""
            )

            # Служебные разделы / акции
            if "dont_miss_it" in classes:
                continue

            # Только прямая ссылка первого уровня
            link = li.locator(
                ":scope > a"
            ).first

            if link.count() == 0:
                continue

            href = link.get_attribute(
                "href"
            )

            name = clean_text(
                link.inner_text()
            )

            if not href or not name:
                continue

            # Служебные ссылки
            if href.startswith("/catalog/"):
                continue

            url = normalize_url(
                href
            )

            # Защита от дублей
            if url in seen:
                continue

            seen.add(url)

            categories.append({
                "name": name,
                "url": url
            })



        return categories

    except Exception as e:

        log(
            f"❌ Ошибка категорий: {e}"
        )

        return []


# ==========================================================
# ОТКРЫТИЕ КАТЕГОРИИ + ПАГИНАЦИЯ
# ==========================================================

def open_category(page, category):

    try:
        page.goto(
            category["url"],
            wait_until="domcontentloaded",
            timeout=60000
        )

        time.sleep(2)

    except Exception as e:
        log(
            f"❌ Ошибка загрузки категории "
            f"{category['name']}: {e}"
        )
        return ""

    while True:

        product_locator = page.locator(
            "tr.itemPosition.simple"
        )

        before_count = product_locator.count()

        more_button = page.locator(
            "#moreBtn"
        )

        if more_button.count() == 0:
            break

        try:
            if not more_button.is_visible():
                break
        except Exception:
            break

        try:
            more_button.scroll_into_view_if_needed()

            more_button.click(
                timeout=10000
            )

        except Exception:
            break

        try:

            page.wait_for_function(
                """
                (before) => {
                    return document.querySelectorAll(
                        'tr.itemPosition.simple'
                    ).length > before;
                }
                """,
                arg=before_count,
                timeout=15000
            )

        except Exception:
            time.sleep(2)

        after_count = product_locator.count()

        if after_count <= before_count:
            break

    return page.content()





# ==========================================================
# АВТОРИЗАЦИЯ
# ==========================================================
def login(page):

    log("🔐 Dellta: открываем сайт")

    try:
        page.goto(
            BASE_URL + "/",
            wait_until="domcontentloaded",
            timeout=60000
        )
    except Exception as e:
        log(f"❌ Dellta: ошибка открытия сайта: {e}")
        return False

    time.sleep(5)

    # ======================================================
    # ПРОВЕРЯЕМ ADM.TOOLS
    # ======================================================

    try:
        body_text = clean_text(
            page.locator("body").inner_text()
        )
    except Exception:
        body_text = ""

    if (
        "Сторінка захищена" in body_text
        or "захищено adm.tools" in body_text
        or "Продовжити" in body_text
    ):

        log("🛡️ Dellta: требуется проверка сайта")

        # Сообщаем bot.py, что пользователю нужно
        # открыть браузер и пройти проверку.
        save_status(
            True,
            0,
            USER,
            str(OUTPUT_FILE),
            verification_required=True
        )

        protection_passed = False

        # Ждём ручного прохождения проверки
        # максимум 5 минут.

        for _ in range(60):

            time.sleep(5)

            try:
                current_text = clean_text(
                    page.locator("body").inner_text()
                )
            except Exception:
                current_text = ""

            if (
                "Сторінка захищена" not in current_text
                and "захищено adm.tools" not in current_text
                and "Продовжити" not in current_text
            ):

                protection_passed = True

                log("✅ Dellta: проверка сайта пройдена")

                save_status(
                    True,
                    0,
                    USER,
                    str(OUTPUT_FILE),
                    verification_required=False
                )

                break

        if not protection_passed:

            log("❌ Dellta: время проверки истекло")

            save_status(
                False,
                0,
                USER,
                str(OUTPUT_FILE),
                verification_required=False
            )

            return False

        time.sleep(3)

    else:

        # Проверка не требуется
        save_status(
            True,
            0,
            USER,
            str(OUTPUT_FILE),
            verification_required=False
        )

    # ======================================================
    # ПРОВЕРЯЕМ, НЕ АВТОРИЗОВАНЫ ЛИ УЖЕ
    # ======================================================

    logout_button = page.get_by_text(
        "Вийти",
        exact=True
    )

    if logout_button.count():

        log("✅ Dellta: пользователь уже авторизован")
        return True

    profile_button = page.get_by_text(
        "Профіль",
        exact=True
    )

    if profile_button.count():

        log("✅ Dellta: пользователь уже авторизован")
        return True

    # ======================================================
    # ИЩЕМ КНОПКУ ВХОДА
    # ======================================================

    enter_button = page.locator(
        "#a-enter"
    )

    if not enter_button.count():

        log("❌ Dellta: кнопка входа не найдена")
        return False

    try:

        enter_button.click(
            timeout=15000
        )

    except Exception as e:

        log(
            f"❌ Dellta: ошибка открытия формы: {e}"
        )

        return False

    # ======================================================
    # ЖДЁМ ФОРМУ
    # ======================================================

    try:

        page.locator(
            "#login-form"
        ).wait_for(
            state="visible",
            timeout=15000
        )

    except PlaywrightTimeoutError:

        log(
            "❌ Dellta: форма авторизации не появилась"
        )

        return False

    email_input = page.locator(
        "input[name='email_auth']"
    )

    password_input = page.locator(
        "input[name='pass_auth']"
    )

    if not email_input.count():

        log(
            "❌ Dellta: поле email не найдено"
        )

        return False

    if not password_input.count():

        log(
            "❌ Dellta: поле пароля не найдено"
        )

        return False

    # ======================================================
    # ВВОДИМ ДАННЫЕ
    # ======================================================

    try:

        email_input.fill(
            EMAIL
        )

        password_input.fill(
            PASSWORD
        )

    except Exception as e:

        log(
            f"❌ Dellta: ошибка заполнения формы: {e}"
        )

        return False

    # ======================================================
    # ВХОД
    # ======================================================

    login_button = page.locator(
        "#login-form button.modalButton"
    )

    if not login_button.count():

        log(
            "❌ Dellta: кнопка входа не найдена"
        )

        return False

    try:

        login_button.click(
            timeout=15000
        )

    except Exception as e:

        log(
            f"❌ Dellta: ошибка кнопки входа: {e}"
        )

        return False

    # ======================================================
    # ЖДЁМ РЕЗУЛЬТАТ
    # ======================================================

    time.sleep(5)

    try:

        if page.locator(
            "#login-form"
        ).is_visible():

            log(
                "❌ Dellta: форма авторизации всё ещё открыта"
            )

            return False

    except Exception:
        pass

    time.sleep(2)

    try:

        current_body = clean_text(
            page.locator(
                "body"
            ).inner_text()
        )

    except Exception:

        current_body = ""

    if (
        "Вийти" in current_body
        or "Профіль" in current_body
    ):

        log(
            "✅ Dellta: авторизация выполнена"
        )

        return True

    log(
        "❌ Dellta: авторизацию подтвердить не удалось"
    )

    return False


# ==========================================================
# ОСНОВНОЙ ПАРСЕР
# ==========================================================

def run_parser():

    if is_locked():
        return

    set_lock(True)

    try:

        # --------------------------------------------------
        # START STATUS
        # --------------------------------------------------

        save_status(
            True,
            0,
            USER,
            str(OUTPUT_FILE)
        )

        log(
            "🚀 Запуск парсера Dellta"
        )

        # --------------------------------------------------
        # EXCEL
        # --------------------------------------------------

        reset_excel()

        # --------------------------------------------------
        # PLAYWRIGHT
        # --------------------------------------------------

        with sync_playwright() as p:
        
            context = p.firefox.launch_persistent_context(
                user_data_dir="/app/output/dellta_firefox",
                headless=False,
                viewport={"width": 1440, "height": 900},
                locale="uk-UA",
                timezone_id="Europe/Kyiv",
                user_agent=(
                    "Mozilla/5.0 "
                    "(X11; Linux x86_64; rv:128.0) "
                    "Gecko/20100101 Firefox/128.0"
                )
            )
        
            if context.pages:
                page = context.pages[0]
            else:
                page = context.new_page()


            # ------------------------------------------------
            # АВТОРИЗАЦИЯ
            # ------------------------------------------------

            if not login(page):

                log("❌ LOGIN FAILED")

                reset_excel()

                save_status(
                    False,
                    0,
                    USER,
                    str(OUTPUT_FILE)
                )


                context.close()

                return

            # ------------------------------------------------
            # КАТЕГОРИИ
            # ------------------------------------------------

            categories = get_categories(
                page
            )

            log(
                f"📂 Категорий получено: {len(categories)}"
            )

            if not categories:

                reset_excel()

                save_status(
                    False,
                    0,
                    USER,
                    str(OUTPUT_FILE)
                )


                context.close()

                return


            # ------------------------------------------------
            # LIMIT
            # ------------------------------------------------

            if CATEGORY_LIMIT:

                categories = categories[
                    :CATEGORY_LIMIT
                ]

            total_categories = len(
                categories
            )

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


                    progress = int(
                        index
                        / total_categories
                        * 100
                    )

                    save_status(
                        True,
                        progress,
                        USER,
                        str(OUTPUT_FILE)
                    )


                except Exception as e:

                    log(
                        f"❌ Ошибка категории {category['name']}: {e}"
                    )
                
                    continue


            # ------------------------------------------------
            # ЕСЛИ ВСЕ КАТЕГОРИИ УПАЛИ
            # ------------------------------------------------

            if successful_categories == 0:

                reset_excel()

                save_status(
                    False,
                    0,
                    USER,
                    str(OUTPUT_FILE)
                )


                context.close()

                return


            # ------------------------------------------------
            # FINISH
            # ------------------------------------------------
            context.close()

        save_status(
            False,
            100,
            USER,
            str(OUTPUT_FILE)
        )

        log(
            "🎉 DELLTA ЗАВЕРШЕН"
        )

    except Exception as e:

        # --------------------------------------------------
        # КРИТИЧЕСКАЯ ОШИБКА
        # --------------------------------------------------

        reset_excel()

        log(
            f"❌ Критическая ошибка: {e}"
        )

        save_status(
            False,
            0,
            USER,
            str(OUTPUT_FILE)
        )

    finally:

        set_lock(False)


# ==========================================================
# START
# ==========================================================

if __name__ == "__main__":
    run_parser()
