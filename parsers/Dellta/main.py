
import os
import sys
import json
import re
import time
import requests

from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook


# ==========================================================
# USER
# ==========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"


# ==========================================================
# SETTINGS
# ==========================================================

print("🔥 DELLTA LIFE PARSER")

BASE = "https://b2b.delltalife.com"

# Для теста можно поставить 2
CATEGORY_LIMIT = 2
# CATEGORY_LIMIT = None

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "123456"

OUTPUT_DIR = os.path.abspath("output/Dellta")
FILE_PATH = os.path.join(OUTPUT_DIR, "Dellta_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")


# ==========================================================
# HTTP
# ==========================================================

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,*/*;q=0.8"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en-US;q=0.7,en;q=0.6",
    "Connection": "keep-alive",
}

session = requests.Session()
session.headers.update(HEADERS)


# ==========================================================
# STATUS
# ==========================================================

def save_status(
    running=False,
    progress=0,
    user="",
    file_path=""
):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path,
    }

    tmp = STATUS_PATH + ".tmp"

    try:
        with open(tmp, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )

        os.replace(tmp, STATUS_PATH)

    except Exception as e:
        print(f"⚠ Не удалось сохранить status.json: {e}")


# ==========================================================
# HELPERS
# ==========================================================

def clean(text):
    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        text
    ).strip()


def absolute_url(url):
    if not url:
        return ""

    if url.startswith("http://") or url.startswith("https://"):
        return url

    if url.startswith("/"):
        return BASE.rstrip("/") + url

    return BASE.rstrip("/") + "/" + url.lstrip("/")


# ==========================================================
# LOGIN
# ==========================================================

def login():
    print("🔐 Авторизация Dellta...")

    login_page = BASE + "/"

    try:
        print(f"🌐 GET {login_page}")

        r = session.get(
            login_page,
            timeout=(15, 60),
            allow_redirects=True
        )

        print(f"🌐 HTTP: {r.status_code}")
        print(f"🌐 Final URL: {r.url}")

    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка GET главной страницы: {e}")
        return False

    # ------------------------------------------------------
    # 429
    # ------------------------------------------------------

    if r.status_code == 429:
        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        title = clean(
            soup.title.get_text()
            if soup.title
            else ""
        )

        print("❌ Dellta вернул HTTP 429")

        if title:
            print(f"❌ Response title: {title}")

        text = clean(soup.get_text(" ", strip=True))

        if text:
            print(f"❌ Ответ: {text[:500]}")

        print()
        print(
            "⚠ Сервер Dellta не отдал страницу с модалкой "
            "авторизации."
        )
        print(
            "⚠ Поэтому заполнить #login-form сейчас невозможно."
        )

        return False

    # ------------------------------------------------------
    # Другие HTTP ошибки
    # ------------------------------------------------------

    if r.status_code >= 400:
        print(
            f"❌ Ошибка открытия Dellta: HTTP {r.status_code}"
        )
        return False

    # ------------------------------------------------------
    # Парсим страницу
    # ------------------------------------------------------

    soup = BeautifulSoup(
        r.text,
        "html.parser"
    )

    # Реальная форма из модального окна
    form = soup.select_one("#login-form")

    if not form:
        # Дополнительный поиск по action
        form = soup.select_one(
            'form[action="/themes/default/ajax/login.php"]'
        )

    if not form:
        print("❌ #login-form не найден")

        title = clean(
            soup.title.get_text()
            if soup.title
            else ""
        )

        if title:
            print(f"📄 Response title: {title}")

        return False

    print("✅ Модалка авторизации найдена")

    # ------------------------------------------------------
    # Получаем action
    # ------------------------------------------------------

    action = form.get("action")

    if not action:
        action = "/themes/default/ajax/login.php"

    action = absolute_url(action)

    print(f"🔑 LOGIN URL: {action}")

    # ------------------------------------------------------
    # Собираем hidden-поля формы
    # ------------------------------------------------------

    payload = {}

    for inp in form.select("input"):
        name = inp.get("name")

        if not name:
            continue

        input_type = (
            inp.get("type", "text")
            .lower()
        )

        # Пароли/логины зададим ниже сами
        if name in ("email_auth", "pass_auth"):
            continue

        # Не отправляем reCAPTCHA регистрации
        if "captcha" in name.lower():
            continue

        if input_type in (
            "submit",
            "button",
            "reset",
        ):
            continue

        payload[name] = inp.get(
            "value",
            ""
        )

    # ------------------------------------------------------
    # Реальные поля Dellta
    # ------------------------------------------------------

    payload["email_auth"] = EMAIL
    payload["pass_auth"] = PASSWORD

    print("📨 Отправка формы авторизации...")

    try:
        r2 = session.post(
            action,
            data=payload,
            headers={
                "X-Requested-With": "XMLHttpRequest",
                "Referer": r.url,
                "Origin": BASE,
                "Accept": "*/*",
            },
            timeout=(15, 60),
            allow_redirects=True,
        )

    except requests.exceptions.RequestException as e:
        print(f"❌ Ошибка POST авторизации: {e}")
        return False

    print(f"🔐 LOGIN HTTP: {r2.status_code}")

    if r2.status_code == 429:
        print("❌ Dellta вернул HTTP 429 при авторизации")
        return False

    if r2.status_code >= 400:
        print(
            f"❌ Ошибка авторизации: HTTP {r2.status_code}"
        )
        return False

    # ------------------------------------------------------
    # Анализ ответа login.php
    # ------------------------------------------------------

    login_text = clean(
        r2.text
    )

    if login_text:
        print(
            f"📄 LOGIN RESPONSE: "
            f"{login_text[:500]}"
        )

    # ------------------------------------------------------
    # Проверяем авторизацию повторным GET
    # ------------------------------------------------------

    print("🔎 Проверяем авторизацию...")

    try:
        check = session.get(
            BASE + "/",
            timeout=(15, 60),
            allow_redirects=True,
        )

    except requests.exceptions.RequestException as e:
        print(
            f"❌ Ошибка проверки авторизации: {e}"
        )
        return False

    print(
        f"🔎 CHECK HTTP: {check.status_code}"
    )

    if check.status_code == 429:
        print(
            "❌ Проверка авторизации получила HTTP 429"
        )
        return False

    check_soup = BeautifulSoup(
        check.text,
        "html.parser"
    )

    # Если форма логина всё ещё присутствует,
    # считаем, что авторизация не прошла.
    still_login = check_soup.select_one(
        "#login-form"
    )

    # Ищем признаки выхода/личного кабинета
    logout = (
        check_soup.select_one(
            'a[href*="logout"]'
        )
        or check_soup.select_one(
            'a[href*="exit"]'
        )
    )

    if still_login and not logout:
        print("❌ Авторизация не подтверждена")
        return False

    print("✅ Dellta: авторизация успешна")

    return True


# ==========================================================
# GET SOUP
# ==========================================================

def get_soup(url):
    try:
        r = session.get(
            url,
            timeout=(15, 60),
            allow_redirects=True,
        )

        if r.status_code == 429:
            print(
                f"❌ HTTP 429: {url}"
            )
            return None

        r.raise_for_status()

        return BeautifulSoup(
            r.text,
            "html.parser"
        )

    except requests.exceptions.RequestException as e:
        print(
            f"❌ GET error: {url} → {e}"
        )
        return None


# ==========================================================
# CATEGORIES
# ==========================================================

def get_categories():
    soup = get_soup(BASE)

    if not soup:
        return []

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
        href = a.get("href")

        if not href:
            continue

        href = absolute_url(href)

        if href not in categories:
            categories.append(href)

    print(
        f"📂 Найдено категорий: "
        f"{len(categories)}"
    )

    return categories


# ==========================================================
# PAGINATION
# ==========================================================

def get_last_page(soup):
    pages = []

    for a in soup.select(
        ".pagination .page-link[pn]"
    ):
        pn = a.get("pn")

        if pn and pn.isdigit():
            pages.append(
                int(pn)
            )

    return max(pages) if pages else 1


# ==========================================================
# PRODUCT PARSER
# ==========================================================

def parse_category(cat_url):
    all_items = []

    print(
        f"📂 CATEGORY: {cat_url}"
    )

    first_page = get_soup(
        cat_url
    )

    if not first_page:
        return all_items

    last_page = get_last_page(
        first_page
    )

    print(
        f"📄 Страниц: {last_page}"
    )

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:
            soup = first_page

        else:
            page_url = (
                cat_url.rstrip("/")
                + f"/page={page}/"
            )

            soup = get_soup(
                page_url
            )

            if not soup:
                continue

        cards = soup.select(
            "tr.itemPosition"
        )

        print(
            f"   Страница {page}: "
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
            # PRICE
            #
            # Берём именно:
            #
            # Комп. ДИЛЕР
            #
            # <tr class="line-1">
            #     <span class="active">
            #         12.80 $
            #     </span>
            #
            # НЕ берём:
            # ОПТ
            # 13.50 $
            # 576.00 ₴
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
            # Добавляем только товар с названием
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
# MAIN PARSER
# ==========================================================

def run_parser():

    print("🚀 Запуск парсера Dellta")

    save_status(
        True,
        0,
        USER,
        FILE_PATH
    )

    # ------------------------------------------------------
    # LOGIN
    # ------------------------------------------------------

    if not login():

        print("❌ LOGIN FAILED")

        save_status(
            False,
            0,
            USER,
            FILE_PATH
        )

        return

    # ------------------------------------------------------
    # CATEGORIES
    # ------------------------------------------------------

    cats = get_categories()

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

    # ------------------------------------------------------
    # LIMIT
    # ------------------------------------------------------

    if CATEGORY_LIMIT:
        cats = cats[:CATEGORY_LIMIT]

        print(
            f"⚠ Тестовый лимит категорий: "
            f"{CATEGORY_LIMIT}"
        )

    total = len(cats)

    print(
        f"🚀 Будет обработано категорий: "
        f"{total}"
    )

    # ------------------------------------------------------
    # EXCEL
    # ------------------------------------------------------

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

    # ------------------------------------------------------
    # PARSE
    # ------------------------------------------------------

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
            cat
        )

        print(
            f"   Получено товаров: "
            f"{len(items)}"
        )

        for item in items:
            ws.append(item)

        # Небольшая пауза между категориями
        time.sleep(0.3)

    # ------------------------------------------------------
    # SAVE
    # ------------------------------------------------------

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    tmp = FILE_PATH + ".tmp"

    try:

        wb.save(tmp)

        os.replace(
            tmp,
            FILE_PATH
        )

    except Exception as e:

        print(
            f"❌ Ошибка сохранения Excel: "
            f"{e}"
        )

        try:
            if os.path.exists(tmp):
                os.remove(tmp)
        except Exception:
            pass

        save_status(
            False,
            0,
            USER,
            FILE_PATH
        )

        return

    # ------------------------------------------------------
    # DONE
    # ------------------------------------------------------

    save_status(
        False,
        100,
        USER,
        FILE_PATH
    )

    print()
    print("================================")
    print("✅ DELLTA DONE")
    print(
        f"📄 FILE: {FILE_PATH}"
    )
    print("================================")


# ==========================================================
# START
# ==========================================================

if __name__ == "__main__":
    run_parser()
