import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys

sys.stdout.reconfigure(line_buffering=True)
sys.stderr.reconfigure(line_buffering=True)

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 Харьковская 4421-4422 Jmax")

BASE = "https://www.jmaxtvshop.com.ua"

EMAIL = "angelinatitor@gmail.com"
PASSWORD = "18022021"

# =========================
# ⚙️ SWITCH
# =========================
#CATEGORY_LIMIT = 1
CATEGORY_LIMIT = None

OUTPUT_DIR = os.path.abspath("output/4421-4422_Jmax")
FILE_PATH = os.path.join(OUTPUT_DIR, "4421-4422_Jmax_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
    "Referer": BASE
}

session = requests.Session()
session.headers.update(HEADERS)

def is_locked():

    if not os.path.exists(LOCK_FILE):
        return False

    try:
        age = time.time() - os.path.getmtime(LOCK_FILE)

        if age > 3600:
            os.remove(LOCK_FILE)
            return False

        return True

    except:
        return False


def set_lock(state):

    if state:

        with open(LOCK_FILE, "w", encoding="utf-8") as f:
            f.write(str(time.time()))

    else:

        if os.path.exists(LOCK_FILE):
            os.remove(LOCK_FILE)
        
# =========================
# STATUS
# =========================
def save_status(running=False, progress=0, user="", file_path=""):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    data = {
        "running": running,
        "progress": progress,
        "user": user,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "file_path": file_path
    }

    tmp = STATUS_PATH + ".tmp"

    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    os.replace(tmp, STATUS_PATH)


# =========================
# HTTP
# =========================
def get_soup(url):

    for _ in range(3):

        try:

            r = session.get(
                url,
                timeout=30,
                allow_redirects=True
            )

            if r.status_code == 200:
                return BeautifulSoup(r.text, "html.parser")

        except Exception as e:
            print(e)

        time.sleep(1)

    return BeautifulSoup("", "html.parser")

def get_product_fast(product_id):

    url = BASE + f"/index.php?route=product/product&product_id={product_id}"

    try:

        r = session.get(
            url,
            timeout=5,
            allow_redirects=True
        )

        if r.status_code != 200:
            return None

        soup = BeautifulSoup(r.text, "html.parser")

        h1 = soup.select_one("h1")

        if not h1:
            return None

        title = clean(h1.get_text())

        if not title:
            return None

        return parse_product_soup(soup, url)

    except Exception as e:

        print(f"❌ ID {product_id}: {e}")
        return None

def login():

    print("🔐 LOGIN...", flush=True)

    login_url = "https://www.jmaxtvshop.com.ua/index.php?route=account/login"

    print(
        f"🔐 LOGIN URL: {login_url}",
        flush=True
    )

    try:

        # ======================================================
        # ТЕСТ ГЛАВНОЙ СТРАНИЦЫ
        # ======================================================

        print(
            "🧪 TEST MAIN PAGE...",
            flush=True
        )

        test = session.get(
            "https://www.jmaxtvshop.com.ua/",
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🧪 MAIN STATUS: {test.status_code} | "
            f"URL: {test.url} | "
            f"HTML: {len(test.text)}",
            flush=True
        )

        # ======================================================
        # ПАУЗА
        # ======================================================

        time.sleep(2)

        # ======================================================
        # ПОЛУЧАЕМ СТРАНИЦУ LOGIN
        # ======================================================

        r = session.get(
            login_url,
            timeout=30,
            allow_redirects=True
        )

        print(
            f"🔐 STATUS: {r.status_code}",
            flush=True
        )

        print(
            f"🔐 FINAL URL: {r.url}",
            flush=True
        )

        print(
            f"🔐 HTML LENGTH: {len(r.text)}",
            flush=True
        )

        # ======================================================
        # 429
        # ======================================================

        if r.status_code == 429:

            print(
                "⚠️ Сайт вернул 429 — слишком много запросов.",
                flush=True
            )

            print(
                "⏳ Ждём 30 секунд и пробуем ещё раз...",
                flush=True
            )

            time.sleep(30)

            r = session.get(
                login_url,
                timeout=30,
                allow_redirects=True
            )

            print(
                f"🔐 RETRY STATUS: {r.status_code}",
                flush=True
            )

            print(
                f"🔐 RETRY URL: {r.url}",
                flush=True
            )

            print(
                f"🔐 RETRY HTML LENGTH: {len(r.text)}",
                flush=True
            )

        # ======================================================
        # ЕСЛИ САЙТ НЕ ОТДАЛ СТРАНИЦУ
        # ======================================================

        if r.status_code != 200:

            print(
                f"❌ Страница логина недоступна: "
                f"HTTP {r.status_code}",
                flush=True
            )

            return False

        # ======================================================
        # HTML
        # ======================================================

        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )

        # ======================================================
        # ИЩЕМ ФОРМУ
        # ======================================================

        form = soup.select_one("form")

        if not form:

            print(
                "❌ LOGIN FORM NOT FOUND",
                flush=True
            )

            forms = soup.find_all("form")

            print(
                f"🔎 Найдено FORM: {len(forms)}",
                flush=True
            )

            return False

        print(
            "✅ LOGIN FORM FOUND",
            flush=True
        )

        # ======================================================
        # ACTION
        # ======================================================

        action = form.get("action")

        if not action:

            action = login_url

        if not action.startswith("http"):

            action = (
                "https://www.jmaxtvshop.com.ua/"
                + action.lstrip("/")
            )

        print(
            f"🔐 FORM ACTION: {action}",
            flush=True
        )

        # ======================================================
        # ПРОВЕРЯЕМ ПОЛЯ
        # ======================================================

        email_input = form.select_one(
            'input[name="email"]'
        )

        password_input = form.select_one(
            'input[name="password"]'
        )

        if not email_input:

            print(
                "❌ Поле email не найдено",
                flush=True
            )

            return False

        if not password_input:

            print(
                "❌ Поле password не найдено",
                flush=True
            )

            return False

        print(
            "✅ Поля email/password найдены",
            flush=True
        )

        # ======================================================
        # LOGIN
        # ======================================================

        payload = {
            "email": EMAIL,
            "password": PASSWORD
        }

        print(
            "🔐 Отправляем LOGIN POST...",
            flush=True
        )

        time.sleep(2)

        r = session.post(
            action,
            data=payload,
            allow_redirects=True,
            timeout=30
        )

        print(
            f"🔐 LOGIN POST STATUS: {r.status_code}",
            flush=True
        )

        print(
            f"🔐 LOGIN FINAL URL: {r.url}",
            flush=True
        )

        print(
            f"🔐 LOGIN HTML LENGTH: {len(r.text)}",
            flush=True
        )

        # ======================================================
        # ПРОВЕРКА УСПЕШНОГО ВХОДА
        # ======================================================

        html_lower = r.text.lower()

        if (
            "account/logout" in html_lower
            or "route=account/logout" in html_lower
            or "account/account" in r.url
        ):

            print(
                "✅ LOGIN OK",
                flush=True
            )

            return True

        # Дополнительная проверка
        if "выйти" in html_lower:

            print(
                "✅ LOGIN OK",
                flush=True
            )

            return True

        print(
            "❌ LOGIN FAILED",
            flush=True
        )

        return False

    except Exception as e:

        print(
            f"❌ LOGIN ERROR: {e}",
            flush=True
        )

        return False



def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""


def get_categories():

    print("")
    print("=" * 70)
    print("🌳 ПОИСК ВСЕХ КАТЕГОРИЙ И СКРЫТЫХ ПОДКАТЕГОРИЙ")
    print("=" * 70)

    categories = []
    seen = set()
    queue = []

    # ==========================================================
    # Нормализация URL
    # ==========================================================

    def normalize_url(url):

        if not url:
            return ""

        url = url.strip()

        if not url:
            return ""

        if url.startswith("#"):
            return ""

        if url.startswith("javascript"):
            return ""

        if not url.startswith("http"):
            url = BASE + "/" + url.lstrip("/")

        if "route=product/category" not in url:
            return ""

        # Убираем пагинацию
        url = re.sub(r"[&?]page=\d+", "", url)

        return url

    # ==========================================================
    # Добавление категории в очередь
    # ==========================================================

    def add_category(url, source=""):

        url = normalize_url(url)

        if not url:
            return

        if url in seen:
            return

        seen.add(url)

        item = {
            "url": url,
            "source": source
        }

        categories.append(item)
        queue.append(url)

        print(
            f"📂 Категория #{len(categories)}: {url}",
            flush=True
        )

    # ==========================================================
    # 1. Главная страница
    # ==========================================================

    soup = get_soup(BASE)

    if not soup or not soup.find_all(True):

        print("❌ Главная страница не загрузилась")

        return []

    print("🏠 Сканируем главную страницу...", flush=True)

    for tag in soup.find_all(True):

        add_category(
            tag.get("href"),
            "MAIN"
        )

        add_category(
            tag.get("data-href"),
            "MAIN"
        )

    # ==========================================================
    # 2. РЕКУРСИВНЫЙ ОБХОД КАТЕГОРИЙ
    # ==========================================================

    index = 0

    while index < len(queue):

        cat_url = queue[index]
        index += 1

        print(
            f"🌳 [{index}/{len(queue)}] Сканируем категорию:",
            cat_url,
            flush=True
        )

        soup = get_soup(cat_url)

        if not soup or not soup.find_all(True):

            print(
                f"⚠️ Не удалось загрузить категорию: {cat_url}",
                flush=True
            )

            continue

        found_before = len(queue)

        # Ищем категории ВО ВСЁМ HTML страницы
        for tag in soup.find_all(True):

            add_category(
                tag.get("href"),
                cat_url
            )

            add_category(
                tag.get("data-href"),
                cat_url
            )

        found_now = len(queue) - found_before

        if found_now:

            print(
                f"   ➕ Найдено новых категорий: {found_now}",
                flush=True
            )

    # ==========================================================
    # ИТОГ
    # ==========================================================

    print("")
    print("=" * 70)
    print(f"🌳 ВСЕГО КАТЕГОРИЙ: {len(categories)}")
    print("=" * 70)

    for i, cat in enumerate(categories, 1):

        print(
            f"{i:04d}. {cat['url']}",
            flush=True
        )

    print("=" * 70)
    print("✅ Полный обход категорий закончен")
    print("=" * 70)

    return categories
   
VISITED_CATEGORIES = set()


def parse_category(cat_url):

    if cat_url in VISITED_CATEGORIES:
        return []

    VISITED_CATEGORIES.add(cat_url)

    result = []

    page = 1

    seen_products = set()

    while True:

        url = f"{cat_url}&page={page}"

        print(
            f"📄 Категория | страница {page}: {url}",
            flush=True
        )

        soup = get_soup(url)

        if not soup or not soup.find_all(True):
            break

        products = []

        for a in soup.find_all("a", href=True):

            href = a.get("href", "").strip()

            if not href:
                continue

            if not href.startswith("http"):
                href = BASE + "/" + href.lstrip("/")

            if "route=product/product" not in href:
                continue

            href = href.split("#")[0]

            if href not in products:
                products.append(href)

        if not products:

            print(
                f"   ⛔ Товаров на странице {page} нет",
                flush=True
            )

            break

        # Проверяем, не повторяет ли сайт предыдущую страницу
        new_products = []

        for href in products:

            if href in seen_products:
                continue

            seen_products.add(href)
            new_products.append(href)

        if not new_products:

            print(
                f"   ⛔ Новых товаров нет — "
                f"останавливаем пагинацию",
                flush=True
            )

            break

        print(
            f"   📦 Найдено новых товаров: {len(new_products)}",
            flush=True
        )

        for href in new_products:

            item = parse_product(href)

            if item and item[1]:

                result.append(item)

            time.sleep(0.05)

        page += 1

    print(
        f"✅ Категория закончена: {len(result)} товаров",
        flush=True
    )

    return result

def parse_search(search_text):

    print(f"🔎 Поиск на сайте: {search_text}")

    url = (
        BASE +
        "/index.php?route=product/search"
        f"&search={search_text}"
        "&description=true"
    )

    soup = get_soup(url)

    collect_product_links(soup)

    products = []

    for a in soup.select(".product-thumb.uni-item a"):

        href = a.get("href", "").strip()

        if not href:
            continue

        if not href.startswith("http"):
            href = BASE + "/" + href.lstrip("/")

        if "route=product/product" not in href:
            continue

        if href not in products:
            products.append(href)

    print(f"🔎 Найдено по запросу {search_text}: {len(products)}")

    result = []

    for href in products:

        result.append(parse_product(href))

        time.sleep(0.05)

    return result



def parse_product_soup(soup, url):

    # =========================
    # TITLE
    # =========================

    title = ""

    h1 = soup.select_one("h1")

    if h1:
        title = clean(h1.get_text())

    # =========================
    # SKU
    # =========================

    sku = ""

    sku_tag = soup.select_one(".product-data__item.model")

    if sku_tag:
        sku = clean(
            sku_tag.get_text().replace("Код товара:", "")
        )

    # =========================
    # PRICE
    # =========================

    price = ""

    p = soup.select_one(".product-page__price")

    if p:
        price = clean(p.get_text())

    # =========================
    # STATUS
    # =========================

    status = ""

    btn = soup.select_one("#button-cart span")

    if btn:
        status = clean(btn.get_text())
    else:
        status = "Нет кнопки"

    return [
        sku,
        title,
        price,
        status,
        url
    ]


def parse_product(url):

    soup = get_soup(url)

    if not soup.select_one("h1"):
        return ["", "", "", "", url]

    return parse_product_soup(soup, url)

# =========================
# MAIN
# =========================
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

        if not login():

            print("❌ Не удалось авторизоваться")

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # ==========================================================
        # СОЗДАЁМ EXCEL
        # ==========================================================

        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])

        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )

        # ==========================================================
        # 🌳 ПОЛНЫЙ ОБХОД ВСЕХ КАТЕГОРИЙ
        # ==========================================================

        cats = get_categories()

        print(
            f"🌳 Получено категорий: {len(cats)}",
            flush=True
        )

        if CATEGORY_LIMIT:
            cats = cats[:CATEGORY_LIMIT]

        all_items = []

        total_categories = len(cats)

        if total_categories == 0:

            print(
                "❌ Категории не найдены",
                flush=True
            )

            save_status(
                False,
                0,
                USER,
                FILE_PATH
            )

            return

        # ==========================================================
        # 📦 ОБХОД КАЖДОЙ КАТЕГОРИИ
        # ==========================================================

        for i, cat in enumerate(cats, 1):

            progress = int(
                (i - 1) / total_categories * 100
            )

            save_status(
                True,
                progress,
                USER,
                FILE_PATH
            )

            print("")
            print("=" * 70)
            print(
                f"📂 КАТЕГОРИЯ {i}/{total_categories}"
            )
            print(
                cat["url"],
                flush=True
            )
            print("=" * 70)

            items = parse_category(
                cat["url"]
            )

            all_items.extend(items)

            print(
                f"📦 Получено товаров из категории: {len(items)}",
                flush=True
            )

            print(
                f"📦 Всего собрано до дедупликации: {len(all_items)}",
                flush=True
            )

        # ==========================================================
        # 🧹 УДАЛЯЕМ ДУБЛИ ТОВАРОВ
        # ==========================================================

        print("")
        print("=" * 70)
        print("🧹 УДАЛЕНИЕ ДУБЛЕЙ")
        print("=" * 70)

        unique_items = []
        seen_products = set()

        for item in all_items:

            sku, title, price, status, url = item

            # Основной ключ — SKU
            key = sku.strip()

            # Если SKU пустой — используем URL
            if not key:
                key = url.strip()

            if not key:
                continue

            if key in seen_products:
                continue

            seen_products.add(key)

            unique_items.append(item)

        items = unique_items

        print(
            f"📦 Было собрано: {len(all_items)}",
            flush=True
        )

        print(
            f"✅ После удаления дублей: {len(items)}",
            flush=True
        )

        print(
            f"🗑️ Удалено дублей: "
            f"{len(all_items) - len(items)}",
            flush=True
        )

        # ==========================================================
        # 💾 ЗАПИСЬ В EXCEL
        # ==========================================================

        print(
            f"📦 Записываем в Excel: {len(items)} товаров",
            flush=True
        )

        total = len(items)

        for i, item in enumerate(items, 1):

            ws.append(item)

            if i % 100 == 0:

                progress = int(
                    i / total * 100
                ) if total else 0

                save_status(
                    True,
                    progress,
                    USER,
                    FILE_PATH
                )

                print(
                    f"💾 Записано: {i}/{total}",
                    flush=True
                )

        # ==========================================================
        # 💾 СОХРАНЕНИЕ ФАЙЛА
        # ==========================================================

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

        print("")
        print("=" * 70)
        print(
            "✅ ГОТОВО. Харьковская 4421-4422 Jmax"
        )
        print(
            f"📊 Всего товаров: {total}"
        )
        print(
            f"📂 Всего обработано категорий: {total_categories}"
        )
        print("=" * 70)

    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
