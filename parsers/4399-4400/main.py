import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

print("🔥 JUMPEX TEST CATEGORY")

BASE = "https://jumpex.com.ua"

CATEGORYS = [
    "https://jumpex.com.ua/instrumenty-i-oborudovanie"
]

OUTPUT_DIR = os.path.abspath("output/4399-4400")
FILE_PATH = os.path.join(
    OUTPUT_DIR,
    "Харьковская_4399-4400_TEST.xlsx"
)

STATUS_PATH = os.path.join(
    OUTPUT_DIR,
    "status.json"
)

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}


def set_status(running):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    with open(STATUS_PATH, "w", encoding="utf-8") as f:
        json.dump(
            {
                "running": running,
                "progress": 0,
                "time": datetime.now().strftime("%d.%m %H:%M")
            },
            f,
            ensure_ascii=False,
            indent=2
        )


def update_progress(percent):
    try:

        with open(STATUS_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)

        data["progress"] = percent

        with open(STATUS_PATH, "w", encoding="utf-8") as f:
            json.dump(
                data,
                f,
                ensure_ascii=False,
                indent=2
            )

    except:
        pass


def clean(text):

    if not text:
        return ""

    return re.sub(
        r"\s+",
        " ",
        text
    ).strip()


def get_soup(url):

    r = requests.get(
        url,
        headers=HEADERS,
        timeout=60
    )

    return BeautifulSoup(
        r.text,
        "html.parser"
    )


def load_products(category_url):

    print("CATEGORY:", category_url)

    links = set()

    start = 0

    while True:

        url = f"{category_url}?start={start}"

        print("LOAD:", url)

        soup = get_soup(url)

        products = soup.select(
            "div.product"
        )

        print(
            "PRODUCTS:",
            len(products)
        )

        if not products:
            break

        before = len(links)

        for product in products:

            a = product.select_one(
                ".name a"
            )

            if not a:
                continue

            href = a.get("href")

            if not href:
                continue

            if href.startswith("/"):
                href = BASE + href

            links.add(href)

        added = len(links) - before

        print("ADDED:", added)

        if added == 0:
            break

        start += 12

        time.sleep(1)

    print("TOTAL LINKS:", len(links))

    return list(links)


def parse_product(url):

    soup = get_soup(url)

    title = ""

    h1 = soup.select_one("h1")

    if h1:
        title = clean(h1.get_text())

    sku = ""

    sku_block = soup.select_one(
        ".prod-ean"
    )

    if sku_block:
        sku = clean(
            sku_block.get_text()
        )
        sku = sku.replace(
            "Артикул:",
            ""
        ).strip()

    price = ""

    price_block = soup.select_one(
        ".prod_price"
    )

    if price_block:
        price = clean(
            price_block.get_text()
        )

    status = ""

    avail = soup.select_one(
        ".avail"
    )

    if avail:
        status = clean(
            avail.get_text()
        )

    return [
        sku,
        title,
        price,
        status,
        url
    ]


def run_parser():

    set_status(True)

    wb = Workbook()
    ws = wb.active

    ws.append([
        "SKU",
        "Title",
        "Price",
        "Status",
        "URL"
    ])

    seen = set()
    total = len(CATEGORYS)

    count = 0

    for i, category in enumerate(CATEGORYS, 1):

        update_progress(
            int(i / total * 100)
        )

        products = load_products(
            category
        )

        for url in products:

            try:

                data = parse_product(
                    url
                )

                sku = data[0]

                if sku in seen:
                    continue

                seen.add(sku)

                ws.append(data)

                count += 1

                print(
                    count,
                    data[1]
                )

            except Exception as e:

                print(
                    "ERROR:",
                    str(e)
                )

    os.makedirs(
        OUTPUT_DIR,
        exist_ok=True
    )

    wb.save(FILE_PATH)

    print("DONE:", count)

    set_status(False)


if __name__ == "__main__":
    run_parser()
