import re
import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin, urlparse, parse_qs

BASE = "https://www.jmaxtvshop.com.ua"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/139.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,uk;q=0.8,en;q=0.7",
}

session = requests.Session()
session.headers.update(HEADERS)


def normalize(url):
    if not url:
        return ""

    url = url.strip()

    if url.startswith("/"):
        url = urljoin(BASE, url)

    if not url.startswith("http"):
        return ""

    url = url.split("#")[0]

    return url


def is_product(url):
    return (
        "route=product/product" in url
        or "/product/" in url
    )


def get(url):
    try:
        print("GET", url)

        r = session.get(
            url,
            timeout=30,
            allow_redirects=True
        )

        print("   HTTP", r.status_code)

        if r.status_code != 200:
            return ""

        return r.text

    except Exception as e:
        print("ERROR:", e)
        return ""


def extract_products(html, source):
    soup = BeautifulSoup(html, "html.parser")

    found = set()

    for a in soup.find_all("a", href=True):

        url = normalize(a.get("href"))

        if not is_product(url):
            continue

        found.add(url)

    print(
        f"   📦 {source}: найдено {len(found)} товаров"
    )

    return found


def main():

    all_products = set()

    sources = [
        ("Главная", BASE),
    ]

    # ==========================================
    # 1. ГЛАВНАЯ
    # ==========================================

    html = get(BASE)

    if html:

        products = extract_products(
            html,
            "Главная"
        )

        all_products.update(products)

    # ==========================================
    # 2. SITEMAP
    # ==========================================

    sitemap_urls = [
        BASE + "/sitemap.xml",
        BASE + "/sitemap_index.xml",
        BASE + "/index.php?route=feed/google_sitemap",
        BASE + "/index.php?route=extension/feed/google_sitemap",
    ]

    print("\n==============================")
    print("SITEMAP")
    print("==============================")

    for sitemap_url in sitemap_urls:

        html = get(sitemap_url)

        if not html:
            continue

        soup = BeautifulSoup(
            html,
            "xml"
        )

        locs = soup.find_all("loc")

        print(
            f"   LOC: {len(locs)}"
        )

        for loc in locs:

            url = normalize(
                loc.get_text(strip=True)
            )

            if not url:
                continue

            # Если sitemap содержит
            # сразу товар
            if is_product(url):

                all_products.add(url)

                continue

            # Если sitemap содержит
            # вложенный sitemap
            if (
                "sitemap" in url.lower()
            ):

                sub_html = get(url)

                if not sub_html:
                    continue

                sub_soup = BeautifulSoup(
                    sub_html,
                    "xml"
                )

                for sub_loc in sub_soup.find_all("loc"):

                    sub_url = normalize(
                        sub_loc.get_text(strip=True)
                    )

                    if is_product(sub_url):

                        all_products.add(
                            sub_url
                        )

    # ==========================================
    # 3. СОХРАНЯЕМ
    # ==========================================

    print("\n==============================")
    print("ИТОГ")
    print("==============================")

    print(
        "ВСЕГО УНИКАЛЬНЫХ ТОВАРОВ:",
        len(all_products)
    )

    with open(
        "jmax_all_products.txt",
        "w",
        encoding="utf-8"
    ) as f:

        for url in sorted(all_products):

            f.write(
                url + "\n"
            )

    print(
        "\n💾 Сохранено:",
        "jmax_all_products.txt"
    )


if __name__ == "__main__":
    main()
