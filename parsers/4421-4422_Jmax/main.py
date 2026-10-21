
# -*- coding: utf-8 -*-
"""
Jmax parser
Excel + status.json + lock.txt
"""

import os
import re
import sys
import json
import time
import logging
from pathlib import Path
from datetime import datetime
from urllib.parse import urljoin, urlparse, urlunparse, parse_qs, urlencode

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook


# ============================================================
# CONFIG
# ============================================================

BASE = "https://www.jmaxtvshop.com.ua"
USER = sys.argv[1] if len(sys.argv) > 1 else "-"

OUTPUT_DIR = Path("output") / "4421-4422_Jmax"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

EXCEL_FILE = OUTPUT_DIR / "4421-4422_Jmax_LIVE.xlsx"
STATUS_FILE = OUTPUT_DIR / "status.json"
LOCK_FILE = OUTPUT_DIR / "lock.txt"

EMAIL = os.getenv("JMAX_EMAIL", "")
PASSWORD = os.getenv("JMAX_PASSWORD", "")

REQUEST_TIMEOUT = 30
RETRIES = 3
REQUEST_DELAY = 0.3
MAX_PAGES_PER_CATEGORY = 300
MAX_CATEGORIES = 1000
MAX_LOCK_AGE = 3600

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:158.0) "
        "Gecko/20100101 Firefox/158.0"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
}

COLUMNS = ["SKU", "TITLE", "PRICE", "STATUS", "URL"]

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s",
    datefmt="%H:%M:%S",
)


# ============================================================
# SESSION
# ============================================================

session = requests.Session()
session.headers.update(HEADERS)


def now():
    return datetime.now().astimezone().isoformat(timespec="seconds")


def save_status(status, message, current=0, total=0, **extra):
    data = {
        "status": status,
        "message": message,
        "current": current,
        "total": total,
        "user": USER,
        "updated_at": now(),
    }
    data.update(extra)

    tmp = STATUS_FILE.with_suffix(".tmp")
    try:
        tmp.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(tmp, STATUS_FILE)
    except Exception:
        logging.exception("Не удалось сохранить status.json")


def acquire_lock():
    if LOCK_FILE.exists():
        try:
            age = time.time() - LOCK_FILE.stat().st_mtime
            if age < MAX_LOCK_AGE:
                logging.error("Парсер уже запущен: %s", LOCK_FILE)
                return False
            LOCK_FILE.unlink(missing_ok=True)
        except OSError:
            logging.exception("Не удалось проверить lock.txt")
            return False

    try:
        LOCK_FILE.write_text(
            json.dumps({"pid": os.getpid(), "started_at": now()}),
            encoding="utf-8",
        )
        return True
    except OSError:
        logging.exception("Не удалось создать lock.txt")
        return False


def release_lock():
    try:
        if LOCK_FILE.exists():
            LOCK_FILE.unlink()
    except OSError:
        logging.exception("Не удалось удалить lock.txt")


# ============================================================
# HTTP
# ============================================================

def normalize_url(url):
    if not url:
        return ""

    url = urljoin(BASE + "/", url.strip())
    parsed = urlparse(url)

    if parsed.netloc.lower() != urlparse(BASE).netloc.lower():
        return ""

    # Удаляем фрагмент, сохраняя параметры страницы.
    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path.rstrip("/") or "/",
        parsed.params,
        parsed.query,
        "",
    ))


def canonical_product_url(url):
    """URL товара без параметров и якорей."""
    url = normalize_url(url)
    if not url:
        return ""

    parsed = urlparse(url)
    return urlunparse((
        parsed.scheme, parsed.netloc, parsed.path.rstrip("/") or "/",
        "", "", ""
    ))


def get_soup(url, params=None):
    for attempt in range(1, RETRIES + 1):
        try:
            response = session.get(
                url,
                params=params,
                timeout=REQUEST_TIMEOUT,
                allow_redirects=True,
            )

            if response.status_code == 429:
                wait = min(60 * attempt, 300)
                logging.warning("HTTP 429. Пауза %s сек.", wait)
                time.sleep(wait)
                continue

            response.raise_for_status()

            if not response.encoding or response.encoding.lower() == "iso-8859-1":
                response.encoding = response.apparent_encoding or "utf-8"

            time.sleep(REQUEST_DELAY)
            return BeautifulSoup(response.text, "html.parser"), response.url

        except requests.RequestException as exc:
            logging.warning(
                "Ошибка запроса %s/%s: %s",
                attempt, RETRIES, exc,
            )
            if attempt < RETRIES:
                time.sleep(attempt * 2)

    return None, url


# ============================================================
# LOGIN
# ============================================================

def login():
    if not EMAIL or not PASSWORD:
        logging.info(
            "JMAX_EMAIL/JMAX_PASSWORD не заданы. Работаем без авторизации."
        )
        return True

    login_url = BASE + "/index.php?route=account/login"
    soup, page_url = get_soup(login_url)

    if soup is None:
        logging.warning("Не удалось открыть страницу входа.")
        return False

    form = soup.select_one("form.jm-login__form")
    if form is None:
        for candidate in soup.select("form"):
            if candidate.select_one('input[name="email"]') and \
               candidate.select_one('input[name="password"]'):
                form = candidate
                break

    if form is None:
        logging.warning("Форма авторизации не найдена.")
        return False

    payload = {}
    for inp in form.select("input[name]"):
        name = inp.get("name")
        typ = (inp.get("type") or "").lower()
        if typ in ("submit", "button", "file"):
            continue
        payload[name] = inp.get("value", "")

    payload["email"] = EMAIL
    payload["password"] = PASSWORD

    action = normalize_url(form.get("action") or page_url)
    method = (form.get("method") or "post").lower()

    try:
        if method == "get":
            response = session.get(
                action, params=payload, timeout=REQUEST_TIMEOUT
            )
        else:
            response = session.post(
                action,
                data=payload,
                headers={"Referer": page_url},
                timeout=REQUEST_TIMEOUT,
            )
        response.raise_for_status()

        check, final_url = get_soup(BASE)
        if check is None:
            return False

        text = check.get_text(" ", strip=True).lower()
        logged_out = (
            "route=account/logout" in str(check)
            or "выйти" in text
            or "выход" in text
        )

        if logged_out:
            logging.info("Авторизация выполнена.")
            return True

        logging.warning(
            "Вход не подтверждён. Продолжаем с текущей сессией."
        )
        return True

    except requests.RequestException:
        logging.exception("Ошибка авторизации.")
        return False


# ============================================================
# EXCEL
# ============================================================

def open_workbook():
    if EXCEL_FILE.exists():
        wb = load_workbook(EXCEL_FILE)
        ws = wb.active

        if ws.max_row == 0:
            ws.append(COLUMNS)
        elif [ws.cell(1, i + 1).value for i in range(len(COLUMNS))] != COLUMNS:
            logging.warning("Заголовки Excel отличаются от ожидаемых.")
        return wb, ws

    wb = Workbook()
    ws = wb.active
    ws.title = "Jmax"
    ws.append(COLUMNS)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:E1"

    for col, width in {
        "A": 16, "B": 65, "C": 16, "D": 22, "E": 90
    }.items():
        ws.column_dimensions[col].width = width

    return wb, ws


def existing_urls(ws):
    result = set()
    for row in ws.iter_rows(min_row=2, min_col=5, max_col=5, values_only=True):
        if row and row[0]:
            result.add(canonical_product_url(str(row[0])))
    return result


def save_workbook(wb):
    tmp = EXCEL_FILE.with_name(EXCEL_FILE.stem + ".tmp.xlsx")
    wb.save(tmp)
    os.replace(tmp, EXCEL_FILE)


# ============================================================
# CATEGORY DISCOVERY
# ============================================================

def looks_like_product_link(anchor):
    classes = " ".join(anchor.get("class", []))
    parent = anchor.find_parent(class_=re.compile(r"product-thumb"))
    return (
        "product-thumb__name" in classes
        or parent is not None
        or anchor.find_parent(class_="product-layout") is not None
    )


def is_probable_category_url(url):
    url = normalize_url(url)
    if not url:
        return False

    parsed = urlparse(url)
    path = parsed.path.lower()

    if not path or path == "/":
        return False

    # Исключаем служебные страницы OpenCart.
    blocked = (
        "/index.php", "/account", "/checkout", "/cart",
        "/search", "/compare", "/wishlist", "/blog",
        "/manufacturer", "/information", "/login",
        "/register", "/contact", "/sitemap",
    )
    if any(part in path for part in blocked):
        return False

    return True


def get_categories():
    soup, _ = get_soup(BASE)
    if soup is None:
        return []

    found = {}
    # Сначала меню: обычные ссылки и data-href.
    for element in soup.select("a[href], [data-href]"):
        if looks_like_product_link(element):
            continue

        href = element.get("href") or element.get("data-href")
        url = normalize_url(href)

        if not is_probable_category_url(url):
            continue

        label = element.get_text(" ", strip=True)
        if not label:
            label = element.get("title", "").strip()

        if not label:
            continue

        path = urlparse(url).path.lower()
        if path in ("/", ""):
            continue

        found[url] = label

    # Добавляем URL корневых категорий, если меню помечено классами.
    for element in soup.select(
        ".menu a[href], .category-menu a[href], "
        ".jm-menu a[href], .dropdown-menu a[href], "
        ".menu [data-href], .category-menu [data-href]"
    ):
        if looks_like_product_link(element):
            continue

        href = element.get("href") or element.get("data-href")
        url = normalize_url(href)
        if url and is_probable_category_url(url):
            found.setdefault(
                url, element.get_text(" ", strip=True) or url
            )

    # Главная категория может не попадать в меню из-за JS-разметки.
    # Её можно указать явно через переменную окружения.
    extra = os.getenv("JMAX_EXTRA_CATEGORIES", "")
    for raw in extra.split(";"):
        raw = raw.strip()
        if not raw:
            continue
        url = normalize_url(raw)
        if url:
            found.setdefault(url, url)

    categories = list(found.items())[:MAX_CATEGORIES]
    logging.info("Найдено URL-кандидатов категорий: %s", len(categories))
    return categories


# ============================================================
# CATEGORY PAGES AND PRODUCT LINKS
# ============================================================

def extract_page_count(soup):
    """
    Например: 'Страница 2 из 5'.
    Возвращает число страниц, если оно найдено.
    """
    for element in soup.select(
        "#jm-inf-more small, .jm-inf__more small, "
        "button.jm-inf__more small"
    ):
        text = element.get_text(" ", strip=True)
        match = re.search(r"страница\s+\d+\s+из\s+(\d+)", text, re.I)
        if match:
            return min(int(match.group(1)), MAX_PAGES_PER_CATEGORY)

    # Иногда пагинация представлена обычными ссылками.
    numbers = []
    for anchor in soup.select("ul.pagination a[href], .pagination a[href]"):
        href = anchor.get("href", "")
        parsed = urlparse(href)
        query = parse_qs(parsed.query)
        for value in query.get("page", []):
            if value.isdigit():
                numbers.append(int(value))

    if numbers:
        return min(max(numbers), MAX_PAGES_PER_CATEGORY)

    return None


def extract_product_links(soup):
    links = {}

    for anchor in soup.select("a.product-thumb__name[href]"):
        url = canonical_product_url(anchor.get("href"))
        title = anchor.get_text(" ", strip=True)

        if url and title:
            links[url] = title

    # Резервный селектор для возможного изменения названия класса.
    if not links:
        for card in soup.select(".product-thumb, .uni-item"):
            anchor = card.select_one("a[href][title], a[href]")
            if not anchor:
                continue

            url = canonical_product_url(anchor.get("href"))
            title = anchor.get_text(" ", strip=True) or anchor.get("title", "")

            if url and title:
                links[url] = title.strip()

    return links


def make_page_url(category_url, page):
    parsed = urlparse(category_url)
    query = parse_qs(parsed.query, keep_blank_values=True)

    if page <= 1:
        query.pop("page", None)
    else:
        query["page"] = [str(page)]

    return urlunparse((
        parsed.scheme,
        parsed.netloc,
        parsed.path,
        parsed.params,
        urlencode(query, doseq=True),
        "",
    ))


def collect_category_products(category_url):
    """
    Собирает ссылки со страниц ?page=1, ?page=2 ...
    Останавливается, когда товары перестают появляться.
    """
    products = {}
    soup1, _ = get_soup(make_page_url(category_url, 1))

    if soup1 is None:
        return products

    products.update(extract_product_links(soup1))
    page_count = extract_page_count(soup1)

    if page_count:
        pages_to_check = page_count
    else:
        # Если счётчик страниц не удалось прочитать,
        # продолжаем до первой пустой страницы.
        pages_to_check = MAX_PAGES_PER_CATEGORY

    for page in range(2, pages_to_check + 1):
        page_url = make_page_url(category_url, page)
        soup, final_url = get_soup(page_url)

        if soup is None:
            logging.warning("Не удалось загрузить страницу %s", page_url)
            if page_count is None:
                break
            continue

        page_products = extract_product_links(soup)

        if not page_products:
            if page_count is None:
                break
            continue

        before = len(products)
        products.update(page_products)

        if len(products) == before and page_count is None:
            break

    return products


# ============================================================
# PRODUCT DETAILS
# ============================================================

def clean_text(element):
    return element.get_text(" ", strip=True) if element else ""


def extract_value(text, pattern):
    match = re.search(pattern, text, re.I)
    return match.group(1).strip() if match else ""


def parse_product(url, listing_title=""):
    soup, final_url = get_soup(url)
    if soup is None:
        return None

    title = clean_text(
        soup.select_one("h1.jm-pp__title")
        or soup.select_one("h1")
    ) or listing_title

    # SKU — код товара, не product_id/data-pid.
    sku = ""
    model = soup.select_one(".product-data__item.model")
    if model:
        model_text = clean_text(model)
        sku = extract_value(
            model_text,
            r"(?:код\s*товара|артикул|модель|код)\s*[:№#]?\s*([A-ZА-ЯЁ0-9._/-]+)"
        )

        if not sku:
            sku = re.sub(
                r"^(?:код\s*товара|артикул|модель|код)\s*[:№#]?\s*",
                "",
                model_text,
                flags=re.I,
            ).strip()

    # Дополнительные селекторы на случай вариаций карточки.
    if not sku:
        for selector in (
            "[itemprop='sku']",
            ".product-model",
            ".product-info .model",
            ".product-data__item",
        ):
            for element in soup.select(selector):
                text = clean_text(element)
                value = extract_value(
                    text,
                    r"(?:код\s*товара|артикул|модель|код)\s*[:№#]?\s*([A-ZА-ЯЁ0-9._/-]+)"
                )
                if value:
                    sku = value
                    break
            if sku:
                break

    price_element = (
        soup.select_one(".product-page__price")
        or soup.select_one(".jm-pp__price")
        or soup.select_one("[itemprop='price']")
        or soup.select_one(".price")
    )

    price = ""
    if price_element:
        price = clean_text(price_element)

        # Если цена хранится только в атрибуте.
        if not price:
            price = (
                price_element.get("content")
                or price_element.get("data-price")
                or ""
            )

    stock_element = (
        soup.select_one(".jm-buy__stock")
        or soup.select_one(".product-data__item.stock")
        or soup.select_one(".stock")
    )
    status = clean_text(stock_element)

    if not status:
        page_text = soup.get_text(" ", strip=True)
        match = re.search(
            r"(В наличии\s*:?\s*\d*\s*шт\.?|Нет в наличии|Под заказ|"
            r"Закончился|Немає в наявності)",
            page_text,
            re.I,
        )
        status = match.group(1).strip() if match else ""

    return [
        sku,
        title,
        price,
        status,
        canonical_product_url(final_url or url),
    ]


# ============================================================
# RUN
# ============================================================

def run_parser():
    if not acquire_lock():
        return

    wb = None

    try:
        save_status("running", "Подготовка парсера")
        login()

        wb, ws = open_workbook()
        known = existing_urls(ws)

        categories = get_categories()

        # Если меню не удалось распознать, можно задать стартовую категорию
        # переменной JMAX_EXTRA_CATEGORIES.
        if not categories:
            save_status(
                "error",
                "Не найдены категории. Проверь меню сайта или JMAX_EXTRA_CATEGORIES.",
            )
            return

        all_products = {}

        for index, (category_url, category_name) in enumerate(categories, 1):
            save_status(
                "running",
                f"Сбор ссылок: {category_name}",
                index,
                len(categories),
                category_url=category_url,
            )
            logging.info(
                "[Категория %s/%s] %s",
                index, len(categories), category_url,
            )

            try:
                links = collect_category_products(category_url)
                all_products.update(links)
                logging.info("Ссылок в категории: %s", len(links))
            except Exception:
                logging.exception("Ошибка категории: %s", category_url)

        # Сохраняем найденные ссылки без дубликатов.
        urls = list(all_products.items())
        total = len(urls)
        logging.info("Всего уникальных товаров: %s", total)

        if total == 0:
            save_status(
                "error",
                "Товары не найдены. Возможно, сайт изменил HTML или требует входа.",
                0,
                0,
            )
            return

        processed = 0
        added = 0
        errors = 0

        for url, listing_title in urls:
            processed += 1

            if url in known:
                continue

            save_status(
                "running",
                f"Обработка товара {processed} из {total}",
                processed,
                total,
                added=added,
                errors=errors,
                url=url,
            )

            try:
                row = parse_product(url, listing_title)
                if row and row[1]:
                    ws.append(row)
                    known.add(url)
                    added += 1

                    # Сохраняем прогресс, чтобы не потерять результат.
                    if added % 10 == 0:
                        save_workbook(wb)

                else:
                    errors += 1

            except Exception:
                errors += 1
                logging.exception("Ошибка карточки: %s", url)

        save_workbook(wb)

        save_status(
            "completed",
            "Парсер завершил работу",
            total,
            total,
            added=added,
            errors=errors,
            excel=str(EXCEL_FILE),
        )

        logging.info(
            "Готово. Новых товаров: %s; ошибок: %s; файл: %s",
            added, errors, EXCEL_FILE,
        )

    except Exception as exc:
        logging.exception("Критическая ошибка парсера")
        save_status("error", str(exc))

    finally:
        release_lock()


if __name__ == "__main__":
    run_parser()

