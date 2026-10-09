
import os
import sys
import re
import json
import time
from datetime import datetime
from urllib.parse import urljoin, urlparse, parse_qs, urlencode, urlunparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment


# ==========================================================
# D-TOP ELECTRONICS — PARSER
# ==========================================================

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 D-Top Electronics — новый парсер")

BASE = "https://www.dtopelectronics.com.ua"
LOGIN_URL = BASE + "/index.php?route=account/login"
ACCOUNT_URL = BASE + "/index.php?route=account/account"

# Прежние учётные данные
EMAIL = "angelinatitor@gmail.com"
PASSWORD = "380931937922"

# None = все категории
# CATEGORY_LIMIT = 1
CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/D-Top")
FILE_PATH = os.path.join(OUTPUT_DIR, "D-Top_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
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
    except OSError:
        return False


def set_lock(state):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    if state:
        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))
    elif os.path.exists(LOCK_FILE):
        try:
            os.remove(LOCK_FILE)
        except OSError:
            pass


# ==========================================================
# STATUS
# ==========================================================

def save_status(running=False, progress=0, user="", file_path="",
                error=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path,
    }

    if error:
        data["error"] = error

    tmp = STATUS_PATH + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, STATUS_PATH)


# ==========================================================
# HELPERS
# ==========================================================

def clean(value):
    return re.sub(r"\s+", " ", value or "").strip()


def absolute_url(href, current_url=BASE):
    if not href:
        return ""

    href = href.strip()

    if href.startswith(("javascript:", "mailto:", "tel:", "#")):
        return ""

    return urljoin(current_url, href)


def normalize_category_url(href, current_url=BASE):
    """
    Приводит ссылки категорий к единому виду.
    Параметры пагинации не входят в адрес категории.
    """
    url = absolute_url(href, current_url)

    if not url:
        return ""

    parsed = urlparse(url)

    if parsed.netloc.lower() != urlparse(BASE).netloc.lower():
        return ""

    params = parse_qs(parsed.query)

    if params.get("route", [""])[0] != "product/category":
        return ""

    path = params.get("path", [""])[0]

    if not path:
        return ""

    query = urlencode({
        "route": "product/category",
        "path": path,
    })

    return urlunparse((
        urlparse(BASE).scheme,
        urlparse(BASE).netloc,
        "/index.php",
        "",
        query,
        "",
    ))


def make_page_url(category_url, page):
    parsed = urlparse(category_url)
    params = parse_qs(parsed.query)

    params["page"] = [str(page)]

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        urlencode(params, doseq=True),
        "",
    ))


def get_soup(url, retries=3):
    for attempt in range(1, retries + 1):
        try:
            response = session.get(url, timeout=40)
            response.raise_for_status()

            if not response.encoding or response.encoding.lower() == "iso-8859-1":
                response.encoding = response.apparent_encoding

            return BeautifulSoup(response.text, "html.parser")

        except requests.RequestException as exc:
            print(
                f"⚠ Ошибка загрузки ({attempt}/{retries}): "
                f"{url} — {exc}"
            )

            if attempt < retries:
                time.sleep(attempt)

    return None


# ==========================================================
# LOGIN
# ==========================================================



def login():
    print("🔐 Авторизация D-Top...", flush=True)

    # 1. Получаем cookies с главной страницы.
    try:
        home = session.get(BASE + "/", timeout=40)
        home.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(
            f"Не удалось открыть главную страницу D-Top: {exc}"
        )

    # 2. Открываем форму входа через ту же сессию.
    soup = get_soup(LOGIN_URL)

    if soup is None:
        raise RuntimeError(
            "Не удалось открыть форму авторизации. "
            "Пароль не отправлялся."
        )

    page_text = clean(soup.get_text(" ", strip=True)).lower()

    lock_messages = (
        "превысила допустимое количество попыток входа",
        "превышено допустимое количество попыток входа",
        "слишком много попыток входа",
        "too many login attempts",
        "too many attempts",
    )

    # Если сайт уже сообщает о блокировке, ничего не отправляем.
    if any(message in page_text for message in lock_messages):
        raise RuntimeError(
            "D-Top заблокировал вход из-за превышения числа попыток. "
            "Новые попытки не выполнялись. Дождитесь разблокировки."
        )

    form = soup.select_one('form[action*="account/login"]')

    if not form:
        raise RuntimeError(
            "Форма авторизации не найдена. Пароль не отправлялся."
        )

    # 3. Подготавливаем поля формы.
    payload = {}

    for inp in form.select("input[name]"):
        name = inp.get("name")
        input_type = (inp.get("type") or "").lower()

        if name and input_type == "hidden":
            payload[name] = inp.get("value", "")

    payload["email"] = EMAIL
    payload["password"] = PASSWORD

    action = absolute_url(form.get("action", ""), LOGIN_URL) or LOGIN_URL

    # 4. Только одна отправка формы за запуск.
    try:
        response = session.post(
            action,
            data=payload,
            headers={"Referer": LOGIN_URL},
            timeout=40,
            allow_redirects=True,
        )
        response.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(
            f"Ошибка отправки формы входа: {exc}. "
            "Автоматических повторных попыток не будет."
        )

    # Проверяем сообщение сайта после единственной попытки.
    response_soup = BeautifulSoup(response.text, "html.parser")
    response_text = clean(
        response_soup.get_text(" ", strip=True)
    ).lower()

    if any(message in response_text for message in lock_messages):
        raise RuntimeError(
            "Сайт сообщил о блокировке входа. "
            "Парсер остановлен без повторных попыток."
        )

    # Проверяем наличие ошибок формы.
    error_el = (
        response_soup.select_one(".alert-danger")
        or response_soup.select_one(".text-danger")
    )

    if error_el:
        error_text = clean(error_el.get_text(" ", strip=True))
        raise RuntimeError(
            f"Сайт отклонил авторизацию: {error_text}. "
            "Повторная отправка пароля не выполнялась."
        )

    # Проверяем доступ к личному кабинету.
    try:
        account = session.get(ACCOUNT_URL, timeout=40)
        account.raise_for_status()
    except requests.RequestException as exc:
        raise RuntimeError(
            f"Не удалось проверить личный кабинет: {exc}"
        )

    account_soup = BeautifulSoup(account.text, "html.parser")
    account_text = clean(
        account_soup.get_text(" ", strip=True)
    ).lower()

    if any(message in account_text for message in lock_messages):
        raise RuntimeError(
            "D-Top сообщил о блокировке входа. "
            "Новые попытки авторизации не выполняются."
        )

    login_form = account_soup.select_one(
        'form[action*="account/login"]'
    )

    if login_form:
        raise RuntimeError(
            "Авторизация не подтверждена. "
            "Пароль повторно отправляться не будет. "
            "Проверьте доступ к аккаунту после окончания блокировки."
        )

    print("✅ Авторизация подтверждена", flush=True)


# ==========================================================
# CATEGORY DISCOVERY
# ==========================================================

def get_categories():
    """
    Собирает категории со всех доступных страниц сайта.
    Поддерживает вложенные пути, например 102_104.
    """
    categories = set()
    checked = set()
    queue = [BASE]

    while queue:
        url = queue.pop(0)

        if url in checked:
            continue

        checked.add(url)

        soup = get_soup(url)

        if soup is None:
            print(f"⚠ Не удалось просканировать категории: {url}")
            continue

        for a in soup.select("a[href]"):
            href = a.get("href", "").strip()

            category_url = normalize_category_url(href, url)

            if not category_url:
                continue

            if category_url not in categories:
                categories.add(category_url)
                queue.append(category_url)

        # Дополнительно ищем категории в ссылках меню.
        for a in soup.select(
            ".menu-module__ul a[href], "
            ".menu-module__children-a[href]"
        ):
            category_url = normalize_category_url(
                a.get("href", ""), url
            )

            if category_url and category_url not in categories:
                categories.add(category_url)
                queue.append(category_url)

        if len(checked) % 25 == 0:
            print(
                f"📂 Сканирование ссылок: проверено {len(checked)}, "
                f"найдено категорий {len(categories)}"
            )

    result = sorted(categories)


    return result


# ==========================================================
# PAGINATION
# ==========================================================

def get_last_page(soup):
    pages = [1]

    for a in soup.select(
        "ul.pagination a[href], .pagination a[href]"
    ):
        href = a.get("href", "")

        try:
            params = parse_qs(urlparse(href).query)
            page_values = params.get("page", [])

            if page_values:
                pages.append(int(page_values[0]))
        except (ValueError, TypeError):
            pass

    return max(pages)


# ==========================================================
# PRODUCT CARD PARSING
# ==========================================================

def parse_product_card(card, category_url):
    title_el = card.select_one(".product-thumb__name")

    if not title_el:
        title_el = card.select_one(
            ".product-thumb__image a[title]"
        )

    if not title_el:
        return None

    title = clean(
        title_el.get("title", "")
        or title_el.get_text(" ", strip=True)
    )

    product_url = absolute_url(
        title_el.get("href", ""),
        category_url,
    )

    if not product_url:
        image_link = card.select_one(".product-thumb__image a[href]")

        if image_link:
            product_url = absolute_url(
                image_link.get("href", ""),
                category_url,
            )

    if not title or not product_url:
        return None

    price_el = card.select_one(".product-thumb__price")

    price = clean(price_el.get_text(" ", strip=True)) if price_el else ""

    status_el = card.select_one(".qty-indicator__text")

    status = clean(status_el.get_text(" ", strip=True)) if status_el else ""

    return {
        "sku": "",
        "title": title,
        "price": price,
        "status": status,
        "url": product_url,
    }


# ==========================================================
# PRODUCT DETAIL — SKU
# ==========================================================

def get_product_sku(product_url):
    soup = get_soup(product_url)

    if soup is None:
        print(f"⚠ Не удалось открыть карточку товара: {product_url}")
        return ""

    sku_el = soup.select_one(
        ".product-data .dtop-code__value"
    )

    if sku_el:
        return clean(sku_el.get_text(" ", strip=True))

    # Запасной вариант: извлечь артикул из блока с подписью «Артикул».
    for item in soup.select(".product-data__item"):
        label = item.select_one(".product-data__item-div")

        if label and clean(label.get_text()).lower() == "артикул":
            value = item.select_one(
                ".dtop-code__value, span"
            )

            if value:
                return clean(value.get_text(" ", strip=True))


    return ""


# ==========================================================
# CATEGORY PARSING
# ==========================================================

def parse_category(category_url):
    products = []
    first_soup = get_soup(category_url)

    if first_soup is None:
        print(f"⚠ Категория недоступна: {category_url}")
        return products

    last_page = get_last_page(first_soup)

    print(f"📄 Страниц в категории: {last_page} — {category_url}")

    for page in range(1, last_page + 1):
        page_url = (
            category_url if page == 1
            else make_page_url(category_url, page)
        )

        soup = first_soup if page == 1 else get_soup(page_url)

        if soup is None:
            print(f"⚠ Пропущена страница: {page_url}")
            continue

        cards = soup.select("div.product-thumb")

        if not cards:
            # Альтернативные варианты разметки каталога.
            cards = soup.select(
                ".product-layout .uni-item, "
                ".product-layout .product-thumb"
            )


        for card in cards:
            item = parse_product_card(card, page_url)

            if item:
                products.append(item)

    return products


# ==========================================================
# EXCEL
# ==========================================================

def save_excel(rows):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "D-Top"

    ws.append(["SKU", "TITLE", "PRICE", "STATUS", "URL"])

    for row in rows:
        ws.append([
            row["sku"],
            row["title"],
            row["price"],
            row["status"],
            row["url"],
        ])

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    widths = {
        "A": 20,
        "B": 65,
        "C": 16,
        "D": 24,
        "E": 80,
    }

    for column, width in widths.items():
        ws.column_dimensions[column].width = width

    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.alignment = Alignment(vertical="center")

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top")

    tmp = FILE_PATH + ".tmp.xlsx"
    wb.save(tmp)
    os.replace(tmp, FILE_PATH)


# ==========================================================
# MAIN
# ==========================================================

def run_parser():
    if is_locked():
        print("⚠ Парсер уже запущен. Выход.")
        return

    set_lock(True)

    try:
        os.makedirs(OUTPUT_DIR, exist_ok=True)

        save_status(True, 0, USER, FILE_PATH)

        # 1. Авторизация
        login()

        # 2. Категории
        categories = get_categories()

        if CATEGORY_LIMIT:
            categories = categories[:CATEGORY_LIMIT]

        if not categories:
            raise RuntimeError(
                "Категории не найдены. Excel не перезаписан."
            )

        # 3. Каталог и карточки товаров
        products_by_url = {}
        total_categories = len(categories)

        for index, category_url in enumerate(categories, 1):
            progress = int((index - 1) / total_categories * 100)

            save_status(True, progress, USER, FILE_PATH)


            category_products = parse_category(category_url)

            for item in category_products:
                # URL товара — основной ключ для удаления дублей.
                key = item["url"]

                if key not in products_by_url:
                    products_by_url[key] = item


            time.sleep(0.15)

        if not products_by_url:
            raise RuntimeError(
                "Товары не найдены. Excel не перезаписан. "
                "Проверьте авторизацию и разметку каталога."
            )

        # 4. Артикулы получаем из карточек товаров.
        items = list(products_by_url.values())
        total_products = len(items)

        for index, item in enumerate(items, 1):
            item["sku"] = get_product_sku(item["url"])

            if index % 20 == 0 or index == total_products:
                progress = 80 + int(index / total_products * 19)

                save_status(True, progress, USER, FILE_PATH)

                print(
                    f"🔎 Карточки: {index}/{total_products}; "
                    f"артикулы получены"
                )

            # Не создаём чрезмерную нагрузку на сайт.
            time.sleep(0.1)

        # 5. Запись Excel только после сбора данных.
        save_excel(items)

        save_status(False, 100, USER, FILE_PATH)

        print("\n✅ D-Top: парсер завершил работу")

    except Exception as exc:
        message = str(exc)

        print(f"\n❌ ОШИБКА D-Top: {message}")

        save_status(
            False,
            0,
            USER,
            FILE_PATH,
            error=message,
        )

    finally:
        set_lock(False)


if __name__ == "__main__":
    run_parser()
