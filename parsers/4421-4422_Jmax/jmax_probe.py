import re
import os
import time
import requests

from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs


BASE = "https://www.jmaxtvshop.com.ua"

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/139.0 Safari/537.36"
    ),
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
}

session = requests.Session()
session.headers.update(HEADERS)


# ==========================================================
# НАСТРОЙКИ
# ==========================================================

OUTPUT_DIR = "/app/parsers/4421-4422_Jmax"

ALL_PRODUCTS_FILE = os.path.join(
    OUTPUT_DIR,
    "jmax_all_products.txt"
)

PRODUCT_IDS_FILE = os.path.join(
    OUTPUT_DIR,
    "jmax_product_ids.txt"
)

CATEGORIES_FILE = os.path.join(
    OUTPUT_DIR,
    "jmax_categories.txt"
)

MISSING_IDS_FILE = os.path.join(
    OUTPUT_DIR,
    "jmax_missing_ids.txt"
)


# ==========================================================
# HTTP
# ==========================================================

def get(url):

    try:

        print("GET", url)

        r = session.get(
            url,
            timeout=30,
            allow_redirects=True
        )

        print(
            "   HTTP",
            r.status_code,
            "FINAL:",
            r.url
        )

        if r.status_code != 200:
            return ""

        return r.text

    except Exception as e:

        print(
            "ERROR:",
            url,
            e
        )

        return ""


# ==========================================================
# URL
# ==========================================================

def normalize(url):

    if not url:
        return ""

    url = url.strip()

    if not url:
        return ""

    if url.startswith("//"):

        url = "https:" + url

    elif url.startswith("/"):

        url = urljoin(
            BASE,
            url
        )

    elif not url.startswith("http"):

        return ""

    url = url.split("#")[0]

    return url


# ==========================================================
# PRODUCT
# ==========================================================

def is_product(url):

    if not url:
        return False

    url_lower = url.lower()

    return (
        "route=product/product" in url_lower
        or "/product/" in url_lower
    )


def get_product_id(url):

    if not url:
        return ""

    parsed = urlparse(url)

    qs = parse_qs(
        parsed.query
    )

    # Обычный OpenCart:
    # ?route=product/product&product_id=123

    values = qs.get(
        "product_id",
        []
    )

    if values:

        value = str(
            values[0]
        ).strip()

        if value.isdigit():

            return value

    # На всякий случай ищем product_id
    m = re.search(
        r"product[_-]?id[=/_-]?(\d+)",
        url,
        flags=re.I
    )

    if m:

        return m.group(1)

    return ""


# ==========================================================
# PRODUCT LINKS
# ==========================================================

def extract_products(
    html,
    source
):

    if not html:
        return set()

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    found = set()

    for a in soup.find_all(
        "a",
        href=True
    ):

        url = normalize(
            a.get("href")
        )

        if not is_product(url):
            continue

        found.add(url)

    print(
        f"   📦 {source}: {len(found)} товаров"
    )

    return found


# ==========================================================
# CATEGORIES
# ==========================================================

def extract_categories(
    html,
    source
):

    if not html:
        return set()

    soup = BeautifulSoup(
        html,
        "html.parser"
    )

    found = set()

    for tag in soup.find_all(
        True
    ):

        for attr in (
            "href",
            "data-href"
        ):

            value = tag.get(
                attr
            )

            if not value:
                continue

            url = normalize(
                value
            )

            if not url:
                continue

            if (
                "route=product/category"
                not in url.lower()
            ):
                continue

            # Убираем page
            url = re.sub(
                r"[&?]page=\d+",
                "",
                url
            )

            found.add(url)

    print(
        f"   📂 {source}: {len(found)} категорий"
    )

    return found


# ==========================================================
# LAST PAGE
# ==========================================================

def get_last_page(
    soup
):

    pages = [1]

    for a in soup.select(
        "ul.pagination a"
    ):

        href = a.get(
            "href",
            ""
        )

        m = re.search(
            r"[?&]page=(\d+)",
            href
        )

        if m:

            pages.append(
                int(
                    m.group(1)
                )
            )

    # Иногда пагинация может быть
    # в другом месте

    for tag in soup.find_all(
        href=True
    ):

        href = tag.get(
            "href",
            ""
        )

        m = re.search(
            r"[?&]page=(\d+)",
            href
        )

        if m:

            pages.append(
                int(
                    m.group(1)
                )
            )

    return max(
        pages
    )


# ==========================================================
# CATEGORY
# ==========================================================

def parse_category(
    category_url
):

    found = set()

    first_html = get(
        category_url
    )

    if not first_html:
        return found

    first_soup = BeautifulSoup(
        first_html,
        "html.parser"
    )

    last_page = get_last_page(
        first_soup
    )

    print(
        f"   📄 Страниц категории: {last_page}"
    )

    for page in range(
        1,
        last_page + 1
    ):

        if page == 1:

            url = category_url

        else:

            separator = (
                "&"
                if "?" in category_url
                else "?"
            )

            url = (
                category_url
                + separator
                + f"page={page}"
            )

        if page != 1:

            html = get(
                url
            )

        else:

            html = first_html

        if not html:
            break

        products = extract_products(
            html,
            f"категория page={page}"
        )

        before = len(
            found
        )

        found.update(
            products
        )

        after = len(
            found
        )

        print(
            f"      +{after - before}, "
            f"всего {after}"
        )

    return found


# ==========================================================
# SITEMAP
# ==========================================================

def parse_sitemap(
    sitemap_url,
    all_products,
    visited_sitemaps
):

    sitemap_url = normalize(
        sitemap_url
    )

    if not sitemap_url:
        return

    if sitemap_url in visited_sitemaps:
        return

    visited_sitemaps.add(
        sitemap_url
    )

    html = get(
        sitemap_url
    )

    if not html:
        return

    soup = BeautifulSoup(
        html,
        "xml"
    )

    locs = soup.find_all(
        "loc"
    )

    print(
        f"   📍 Sitemap LOC: {len(locs)}"
    )

    for loc in locs:

        url = normalize(
            loc.get_text(
                strip=True
            )
        )

        if not url:
            continue

        # Товар непосредственно
        if is_product(url):

            all_products.add(
                url
            )

            continue

        # Вложенный sitemap
        if (
            "sitemap" in url.lower()
        ):

            parse_sitemap(
                url,
                all_products,
                visited_sitemaps
            )


# ==========================================================
# MAIN
# ==========================================================

def main():

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    all_products = set()

    all_categories = set()

    product_ids = {}

    visited_sitemaps = set()


    # ======================================================
    # 1. ГЛАВНАЯ
    # ======================================================

    print()
    print("=" * 70)
    print("1. ГЛАВНАЯ")
    print("=" * 70)

    html = get(
        BASE
    )

    if html:

        products = extract_products(
            html,
            "Главная"
        )

        all_products.update(
            products
        )

        categories = extract_categories(
            html,
            "Главная"
        )

        all_categories.update(
            categories
        )


    # ======================================================
    # 2. SITEMAP
    # ======================================================

    print()
    print("=" * 70)
    print("2. SITEMAP")
    print("=" * 70)

    sitemap_urls = [

        BASE + "/sitemap.xml",

        BASE + "/sitemap_index.xml",

        BASE
        + "/index.php?route=feed/google_sitemap",

        BASE
        + "/index.php?route=extension/feed/google_sitemap",

    ]

    for sitemap_url in sitemap_urls:

        parse_sitemap(
            sitemap_url,
            all_products,
            visited_sitemaps
        )


    # ======================================================
    # 3. КАТЕГОРИИ
    # ======================================================

    print()
    print("=" * 70)
    print("3. КАТЕГОРИИ")
    print("=" * 70)

    print(
        "Всего найдено категорий:",
        len(all_categories)
    )

    # Сохраняем список категорий

    with open(
        CATEGORIES_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        for category in sorted(
            all_categories
        ):

            f.write(
                category
                + "\n"
            )

    # Обходим категории

    category_number = 0

    for category in sorted(
        all_categories
    ):

        category_number += 1

        print()
        print(
            f"📂 КАТЕГОРИЯ "
            f"{category_number}/"
            f"{len(all_categories)}"
        )

        print(
            category
        )

        products = parse_category(
            category
        )

        before = len(
            all_products
        )

        all_products.update(
            products
        )

        after = len(
            all_products
        )

        print(
            f"   ✅ Новых товаров: "
            f"{after - before}"
        )

        print(
            f"   📦 Всего уникальных: "
            f"{after}"
        )


    # ======================================================
    # 4. PRODUCT ID
    # ======================================================

    print()
    print("=" * 70)
    print("4. PRODUCT ID")
    print("=" * 70)

    for url in all_products:

        product_id = get_product_id(
            url
        )

        if product_id:

            product_ids[
                product_id
            ] = url

    print(
        "Товаров с найденным product_id:",
        len(product_ids)
    )


    # ======================================================
    # 5. СОХРАНЯЕМ ВСЕ URL
    # ======================================================

    with open(
        ALL_PRODUCTS_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        for url in sorted(
            all_products
        ):

            f.write(
                url
                + "\n"
            )


    # ======================================================
    # 6. СОХРАНЯЕМ ID
    # ======================================================

    numeric_ids = []

    for product_id in product_ids:

        if product_id.isdigit():

            numeric_ids.append(
                int(product_id)
            )

    numeric_ids.sort()

    with open(
        PRODUCT_IDS_FILE,
        "w",
        encoding="utf-8"
    ) as f:

        for product_id in numeric_ids:

            f.write(
                f"{product_id}\t"
                f"{product_ids[str(product_id)]}\n"
            )


    # ======================================================
    # 7. ПОИСК ПРОПУСКОВ ID
    # ======================================================

    print()
    print("=" * 70)
    print("5. АНАЛИЗ ID")
    print("=" * 70)

    if numeric_ids:

        min_id = min(
            numeric_ids
        )

        max_id = max(
            numeric_ids
        )

        print(
            "MIN ID:",
            min_id
        )

        print(
            "MAX ID:",
            max_id
        )

        expected = set(
            range(
                min_id,
                max_id + 1
            )
        )

        actual = set(
            numeric_ids
        )

        missing = sorted(
            expected - actual
        )

        print(
            "Всего отсутствующих ID:",
            len(missing)
        )

        with open(
            MISSING_IDS_FILE,
            "w",
            encoding="utf-8"
        ) as f:

            for product_id in missing:

                f.write(
                    str(product_id)
                    + "\n"
                )

        if missing:

            print()
            print(
                "Первые пропуски:"
            )

            print(
                missing[:100]
            )

    else:

        print(
            "❌ Числовых product_id "
            "не найдено"
        )


    # ======================================================
    # ИТОГ
    # ======================================================

    print()
    print("=" * 70)
    print("🔥 ИТОГ")
    print("=" * 70)

    print(
        "Категорий:",
        len(all_categories)
    )

    print(
        "Уникальных товаров:",
        len(all_products)
    )

    print(
        "Товаров с product_id:",
        len(product_ids)
    )

    print(
        "Сохранено:",
        ALL_PRODUCTS_FILE
    )

    print(
        "Сохранено:",
        PRODUCT_IDS_FILE
    )

    print(
        "Сохранено:",
        CATEGORIES_FILE
    )

    print(
        "Сохранено:",
        MISSING_IDS_FILE
    )

    print()
    print(
        "✅ РАЗВЕДКА ЗАКОНЧЕНА"
    )


if __name__ == "__main__":

    main()
