import os
import sys
import json
import time
import subprocess
import shutil

from datetime import datetime
from openpyxl import Workbook
from bs4 import BeautifulSoup


# ==========================================================
# USER
# ==========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"


# ==========================================================
# SETTINGS
# ==========================================================

print("🔥 DELLTA LIFE PARSER")

BASE = "https://b2b.delltalife.com"

# Для первого теста можно поставить:
CATEGORY_LIMIT = 2
# CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"

OUTPUT_DIR = os.path.abspath("output/Dellta")
FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Dellta_LIVE.xlsx"
)
STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)


# ==========================================================
# PLAYWRIGHT
# ==========================================================

try:
    from playwright.sync_api import (
        sync_playwright,
        TimeoutError as PlaywrightTimeoutError,
    )
except ImportError:
    print("❌ Playwright не установлен")
    print("❌ Установи: pip install playwright")
    sys.exit(1)


# ==========================================================
# STATUS
# ==========================================================

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
        "file_path": file_path,
    }

    tmp = STATUS_PATH + ".tmp"

    try:
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

    except Exception as e:
        print(
            f"⚠ Ошибка status.json: {e}"
        )


# ==========================================================
# HELPERS
# ==========================================================

def clean(text):
    if not text:
        return ""

    return " ".join(
        text.split()
    ).strip()


def absolute_url(url):
    if not url:
        return ""

    if (
        url.startswith("http://")
        or url.startswith("https://")
    ):
        return url

    if url.startswith("/"):
        return BASE.rstrip("/") + url

    return (
        BASE.rstrip("/")
        + "/"
        + url.lstrip("/")
    )


# ==========================================================
# INSTALL FIREFOX IF NEEDED
# ==========================================================

def ensure_firefox():
    """
    Проверяем наличие Firefox Playwright.
    Если браузера нет — устанавливаем его.
    """

    try:
        from playwright.sync_api import sync_playwright

        with sync_playwright() as p:
            executable = p.firefox.executable_path

        if executable and os.path.exists(executable):
            print("🦊 Firefox Playwright найден")
            return True

    except Exception:
        pass

    print("🦊 Firefox Playwright не найден")
    print("📦 Устанавливаем Firefox...")

    try:
        result = subprocess.run(
            [
                sys.executable,
                "-m",
                "playwright",
                "install",
                "firefox",
            ],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            timeout=300,
        )

        if result.returncode != 0:
            print(
                "❌ Не удалось установить Firefox"
            )
            print(
                result.stdout[-3000:]
            )
            return False

        print("✅ Firefox установлен")

        return True

    except Exception as e:
        print(
            f"❌ Ошибка установки Firefox: {e}"
        )
        return False


# ==========================================================
# LOGIN
# ==========================================================

def login(page):
    print("🔐 Авторизация Dellta...")

    try:
        print(
            f"🌐 GET {BASE}/"
        )

        response = page.goto(
            BASE + "/",
            wait_until="domcontentloaded",
            timeout=90000,
        )

        if response:
            print(
                f"🌐 HTTP: "
                f"{response.status}"
            )

        print(
            f"🌐 URL: {page.url}"
        )

    except Exception as e:
        print(
            f"❌ Ошибка открытия Dellta: {e}"
        )
        return False

    # ------------------------------------------------------
    # Ждём JavaScript / cookie protection
    # ------------------------------------------------------

    print(
        "⏳ Ждём выполнение JavaScript..."
    )

    try:
        page.wait_for_timeout(5000)
    except Exception:
        pass

    # ------------------------------------------------------
    # Проверяем, что страница действительно открылась
    # ------------------------------------------------------

    try:
        body_text = clean(
            page.locator("body").inner_text(
                timeout=10000
            )
        )
    except Exception:
        body_text = ""

    if "Захищена сторінка" in body_text:
        print(
            "❌ adm.tools всё ещё показывает "
            "защищённую страницу"
        )

        print(
            "❌ Firefox не получил доступ к сайту"
        )

        return False

    # ------------------------------------------------------
    # Ищем реальную форму
    # ------------------------------------------------------

    login_form = page.locator(
        "#login-form"
    )

    try:
        login_form.wait_for(
            state="visible",
            timeout=30000
        )

        print(
            "✅ Модалка #login-form найдена"
        )

    except PlaywrightTimeoutError:

        # Иногда форма есть в DOM,
        # но модалка ещё не открыта.
        print(
            "⚠ #login-form не появилась сразу"
        )

        # Ищем кнопку/ссылку входа
        selectors = [
            'a[href="#modalLogin"]',
            'a[href*="modalLogin"]',
            '[data-target="#modalLogin"]',
            '[href="#login"]',
            '.login',
        ]

        opened = False

        for selector in selectors:

            try:
                element = page.locator(
                    selector
                ).first

                if element.count() > 0:
                    if element.is_visible():
                        print(
                            f"🔓 Открываем "
                            f"модалку: {selector}"
                        )

                        element.click(
                            timeout=5000
                        )

                        page.wait_for_timeout(
                            1000
                        )

                        opened = True
                        break

            except Exception:
                continue

        if opened:

            try:
                login_form.wait_for(
                    state="visible",
                    timeout=15000
                )

                print(
                    "✅ Модалка #login-form найдена"
                )

            except Exception:
                pass

    # ------------------------------------------------------
    # Повторно проверяем форму
    # ------------------------------------------------------

    if login_form.count() == 0:

        print(
            "❌ #login-form не найдена"
        )

        try:
            title = clean(
                page.title()
            )

            if title:
                print(
                    f"📄 Page title: {title}"
                )

        except Exception:
            pass

        return False

    # ------------------------------------------------------
    # EMAIL
    # ------------------------------------------------------

    email_input = page.locator(
        '#login-form input[name="email_auth"]'
    )

    password_input = page.locator(
        '#login-form input[name="pass_auth"]'
    )

    if email_input.count() == 0:
        print(
            "❌ Поле email_auth не найдено"
        )
        return False

    if password_input.count() == 0:
        print(
            "❌ Поле pass_auth не найдено"
        )
        return False

    print(
        "✏ Заполняем email..."
    )

    email_input.fill(
        EMAIL
    )

    print(
        "✏ Заполняем пароль..."
    )

    password_input.fill(
        PASSWORD
    )

    # ------------------------------------------------------
    # LOGIN BUTTON
    # ------------------------------------------------------

    login_button = page.locator(
        '#login-form button'
    ).first

    if login_button.count() == 0:

        print(
            "❌ Кнопка входа не найдена"
        )
        return False

    print(
        "🚀 Нажимаем «Увійти»..."
    )

    try:

        login_button.click(
            timeout=15000
        )

    except Exception as e:

        print(
            f"❌ Ошибка нажатия кнопки: {e}"
        )
        return False

    # ------------------------------------------------------
    # Ждём AJAX login
    # ------------------------------------------------------

    page.wait_for_timeout(
        3000
    )

    # ------------------------------------------------------
    # Проверяем ошибки авторизации
    # ------------------------------------------------------

    try:
        current_text = clean(
            page.locator("body").inner_text(
                timeout=10000
            )
        )
    except Exception:
        current_text = ""

    error_words = [
        "Невірний",
        "Неверный",
        "неправильний",
        "неправильный",
        "Помилка",
        "Ошибка",
        "помилка",
        "ошибка",
    ]

    for word in error_words:

        if word in current_text:

            # Это не всегда ошибка логина —
            # слово может быть где-то ещё.
            # Поэтому проверяем саму форму.
            try:
                form_visible = login_form.is_visible()
            except Exception:
                form_visible = False

            if form_visible:

                print(
                    "❌ Dellta не принял авторизацию"
                )

                return False

            break

    # ------------------------------------------------------
    # Проверяем исчезновение login modal
    # ------------------------------------------------------

    try:
        still_visible = login_form.is_visible()
    except Exception:
        still_visible = False

    # ------------------------------------------------------
    # Проверяем признаки авторизации
    # ------------------------------------------------------

    logged_in = False

    auth_selectors = [
        'a[href*="logout"]',
        'a[href*="exit"]',
        '.logout',
        '[href*="profile"]',
        '[href*="cabinet"]',
    ]

    for selector in auth_selectors:

        try:

            element = page.locator(
                selector
            ).first

            if element.count() > 0:
                if element.is_visible():
                    logged_in = True
                    break

        except Exception:
            continue

    if logged_in:
        print(
            "✅ Dellta: авторизация успешна"
        )
        return True

    if not still_visible:

        print(
            "✅ Модалка закрылась"
        )
        print(
            "✅ Dellta: авторизация успешна"
        )

        return True

    # ------------------------------------------------------
    # Последняя проверка через cookies
    # ------------------------------------------------------

    try:
        cookies = page.context.cookies()

        if cookies:

            cookie_names = [
                c.get("name", "")
                for c in cookies
            ]

            print(
                f"🍪 Cookies получено: "
                f"{len(cookie_names)}"
            )

    except Exception:
        pass

    print(
        "❌ Авторизация не подтверждена"
    )

    return False


# ==========================================================
# GET CATEGORIES
# ==========================================================

def get_categories(page):
    print(
        "📂 Получаем категории..."
    )

    try:

        page.goto(
            BASE + "/",
            wait_until="domcontentloaded",
            timeout=90000
        )

        page.wait_for_timeout(
            2000
        )

    except Exception as e:

        print(
            f"❌ Ошибка открытия главной: {e}"
        )

        return []

    html = page.content()

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    categories = []

    container = soup.select_one(
        "div.brandsOnMain"
    )

    if not container:

        print(
            "❌ div.brandsOnMain не найден"
        )

        return []

    for a in container.select(
        "a.COMitem"
    ):

        href = a.get(
            "href"
        )

        if not href:
            continue

        href = absolute_url(
            href
        )

        if href not in categories:

            categories.append(
                href
            )

    print(
        f"📂 Найдено категорий: "
        f"{len(categories)}"
    )

    return categories


# ==========================================================
# LAST PAGE
# ==========================================================

def get_last_page(soup):
    pages = []

    for a in soup.select(
        ".pagination .page-link[pn]"
    ):

        pn = a.get(
            "pn"
        )

        if pn and pn.isdigit():

            pages.append(
                int(pn)
            )

    return (
        max(pages)
        if pages
        else 1
    )


# ==========================================================
# PARSE CATEGORY
# ==========================================================

def parse_category(
    page,
    cat_url
):
    all_items = []

    print(
        f"📂 CATEGORY: {cat_url}"
    )

    try:

        page.goto(
            cat_url,
            wait_until="domcontentloaded",
            timeout=90000
        )

        page.wait_for_timeout(
            1500
        )

    except Exception as e:

        print(
            f"❌ Ошибка категории: {e}"
        )

        return all_items

    # ------------------------------------------------------
    # First page
    # ------------------------------------------------------

    html = page.content()

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    last_page = get_last_page(
        soup
    )

    print(
        f"📄 Страниц: {last_page}"
    )

    # ------------------------------------------------------
    # Pages
    # ------------------------------------------------------

    for page_number in range(
        1,
        last_page + 1
    ):

        if page_number != 1:

            page_url = (
                cat_url.rstrip("/")
                + f"/page={page_number}/"
            )

            try:

                page.goto(
                    page_url,
                    wait_until="domcontentloaded",
                    timeout=90000
                )

                page.wait_for_timeout(
                    1000
                )

            except Exception as e:

                print(
                    f"❌ Ошибка страницы "
                    f"{page_number}: {e}"
                )

                continue

            html = page.content()

            soup = BeautifulSoup(
                html,
                "html.parser"
            )

        cards = soup.select(
            "tr.itemPosition"
        )

        print(
            f"   Страница "
            f"{page_number}: "
            f"{len(cards)} товаров"
        )

        for card in cards:

            sku = ""
            title = ""
            price = ""
            status = ""
            url = ""

            # ------------------------------------------------
            # SKU
            # ------------------------------------------------

            sku_el = card.select_one(
                "td.td_2 .gray"
            )

            if sku_el:

                sku = clean(
                    sku_el.get_text()
                )

            # ------------------------------------------------
            # TITLE + URL
            # ------------------------------------------------

            title_el = card.select_one(
                "td.td_2 a[href]"
            )

            if not title_el:

                title_el = card.select_one(
                    'td.td_2 a[href*="/invertoryi-"]'
                )

            if title_el:

                title = clean(
                    title_el.get_text()
                )

                url = absolute_url(
                    title_el.get(
                        "href",
                        ""
                    )
                )

            # ------------------------------------------------
            # STATUS
            # ------------------------------------------------

            status_el = card.select_one(
                "td.td_2 .are-available"
            )

            if not status_el:

                status_el = card.select_one(
                    "td.td_2 div[class^='are-']"
                )

            if status_el:

                status = clean(
                    status_el.get_text(
                        " ",
                        strip=True
                    )
                )

            # ------------------------------------------------
            # DEALER PRICE
            # ------------------------------------------------
            #
            # Комп. ДИЛЕР
            #
            # <tr class="line-1">
            #     <span class="active">
            #         12.80 $
            #     </span>
            #
            # ------------------------------------------------

            price_el = card.select_one(
                "td.td_3 "
                ".price-table "
                "tr.line-1 "
                "span.active"
            )

            if not price_el:

                price_el = card.select_one(
                    "td.td_3 "
                    "tr.line-1 "
                    "span.active"
                )

            if price_el:

                price = clean(
                    price_el.get_text()
                )

            # ------------------------------------------------
            # ADD
            # ------------------------------------------------

            if not title:
                continue

            all_items.append([
                sku,
                title,
                price,
                status,
                url,
            ])

    return all_items


# ==========================================================
# MAIN
# ==========================================================

def run_parser():

    print(
        "🚀 Запуск парсера Dellta"
    )

    save_status(
        True,
        0,
        USER,
        FILE_PATH
    )

    # ------------------------------------------------------
    # Firefox
    # ------------------------------------------------------

    if not ensure_firefox():

        save_status(
            False,
            0,
            USER,
            FILE_PATH
        )

        return

    # ------------------------------------------------------
    # Browser
    # ------------------------------------------------------

    with sync_playwright() as p:

        browser = None

        try:

            print(
                "🦊 Запускаем Firefox..."
            )

            browser = p.firefox.launch(
                headless=True
            )

            context = browser.new_context(
                viewport={
                    "width": 1366,
                    "height": 900,
                },
                locale="uk-UA",
                timezone_id="Europe/Kyiv",
                user_agent=(
                    "Mozilla/5.0 "
                    "(Windows NT 10.0; Win64; x64; rv:128.0) "
                    "Gecko/20100101 Firefox/128.0"
                ),
            )

            page = context.new_page()

            # ------------------------------------------------
            # LOGIN
            # ------------------------------------------------

            if not login(page):

                print(
                    "❌ LOGIN FAILED"
                )

                save_status(
                    False,
                    0,
                    USER,
                    FILE_PATH
                )

                return

            # ------------------------------------------------
            # CATEGORIES
            # ------------------------------------------------

            cats = get_categories(
                page
            )

            if not cats:

                print(
                    "❌ Категории не найдены"
                )

                save_status(
                    False,
                    0,
                    USER,
                    FILE_PATH
                )

                return

            if CATEGORY_LIMIT:

                cats = cats[
                    :CATEGORY_LIMIT
                ]

                print(
                    f"⚠ Лимит категорий: "
                    f"{CATEGORY_LIMIT}"
                )

            total = len(cats)

            # ------------------------------------------------
            # EXCEL
            # ------------------------------------------------

            wb = Workbook()

            ws = wb.active
            ws.title = "Dellta"

            ws.append([
                "SKU",
                "TITLE",
                "PRICE",
                "STATUS",
                "URL",
            ])

            # ------------------------------------------------
            # PARSE
            # ------------------------------------------------

            seen = set()

            for i, cat in enumerate(
                cats,
                1
            ):

                progress = int(
                    i / total * 100
                )

                save_status(
                    True,
                    progress,
                    USER,
                    FILE_PATH
                )

                print()
                print(
                    f"🔥 [{i}/{total}] "
                    f"{progress}%"
                )

                items = parse_category(
                    page,
                    cat
                )

                print(
                    f"   Получено: "
                    f"{len(items)}"
                )

                for item in items:

                    sku = item[0]
                    url = item[4]

                    key = (
                        sku
                        if sku
                        else url
                    )

                    if key in seen:
                        continue

                    seen.add(
                        key
                    )

                    ws.append(
                        item
                    )

                time.sleep(
                    0.5
                )

            # ------------------------------------------------
            # SAVE
            # ------------------------------------------------

            os.makedirs(
                OUTPUT_DIR,
                exist_ok=True
            )

            tmp = (
                FILE_PATH
                + ".tmp"
            )

            wb.save(
                tmp
            )

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

            print()
            print(
                "================================"
            )
            print(
                "✅ DELLTA DONE"
            )
            print(
                f"📄 FILE: {FILE_PATH}"
            )
            print(
                f"📦 Товаров: {len(seen)}"
            )
            print(
                "================================"
            )

        except Exception as e:

            print(
                f"❌ Критическая ошибка: {e}"
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

        finally:

            try:

                if browser:
                    browser.close()

            except Exception:
                pass


# ==========================================================
# START
# ==========================================================

if __name__ == "__main__":
    run_parser()
