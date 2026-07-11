import os
import json
import re
import time
import requests
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook

import sys

USER = sys.argv[1] if len(sys.argv) > 1 else "-"

print("🔥 LunaBag")

BASE = "https://luna-toys.com.ua"

# =========================
# ⚙️ SWITCH
# =========================

TEST_MODE = True
TEST_LIMIT = 1

# после проверки поменяем:
# TEST_MODE = False

OUTPUT_DIR = os.path.abspath("output/LunaBag")
FILE_PATH = os.path.join(OUTPUT_DIR, "LunaBag_LIVE.xlsx")
STATUS_PATH = os.path.join(OUTPUT_DIR, "status.json")
LOCK_FILE = os.path.join(OUTPUT_DIR, "lock.txt")

HEADERS = {
    "User-Agent": "Mozilla/5.0"
}

session = requests.Session()
session.headers.update(HEADERS)

# Получаем главную страницу
r = session.get(BASE, timeout=30)

# challenge_passed
m = re.search(
    r'document\.cookie\s*=\s*"challenge_passed=([^"]+)',
    r.text
)

if m:
    cookie_value = m.group(1)

    print("🍪 FOUND CHALLENGE COOKIE:", cookie_value[:20])

    session.cookies.set(
        "challenge_passed",
        cookie_value,
        domain="luna-toys.com.ua",
        path="/"
    )

    # перезагрузка после установки cookie
    session.get(
        BASE,
        timeout=30
    )
    
# Загружаем страницу повторно уже с cookie
r = session.get(BASE, timeout=30)
#print(r.text[:1000])

# Ищем CSRF
m = re.search(
    r"GLOBAL_CSRF_TOKEN:\s*'([^']+)'",
    r.text
)

CSRF = m.group(1) if m else ""

print("CSRF:", CSRF)


# Переключаем валюту
try:
    session.get(
        BASE + "/_widget/currency_selector/change/3",
        timeout=30
    )
except:
    pass

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
    
    headers={
        "X-Requested-With": "XMLHttpRequest",
        "X-CSRF-Token": CSRF,
        "Referer": url,
    }

    try:

        r = session.post(
            url,
            data={
                "catalogBuilder": "1"
            },
            headers=headers,
            timeout=30
        )
        
        if r.status_code != 200:
            return BeautifulSoup("", "html.parser")

        data = r.json()

        with open("catalog.json", "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

        #print(data.keys())
        #print(data["response"]["html"].keys())

        html = ""

        if "products" in data["response"]["html"]:
            html += data["response"]["html"]["products"]

        if "pagination" in data["response"]["html"]:
            html += data["response"]["html"]["pagination"]

        return BeautifulSoup(html, "html.parser")


    except Exception as e:
        print(e)
        return BeautifulSoup("", "html.parser")


def clean(t):
    return re.sub(r"\s+", " ", t).strip() if t else ""



# =========================
# PRODUCTS FROM SITEMAP
# =========================
def get_products():

    products = []

    r = session.get(
        BASE + "/sitemap.xml",
        timeout=30
    )

    print("SITEMAP STATUS:", r.status_code)

    print(r.text[:500])

    xml = r.text

    maps = re.findall(
        r"<loc>(.*?)</loc>",
        xml
    )

    for sm in maps:

        if "catalog-sitemap" not in sm:
            continue

        print("📄", sm)

        xml2 = session.get(
            sm,
            timeout=30
        ).text

        urls = re.findall(
            r"<loc>(https://luna-toys\.com\.ua/(?!ru/).*?)</loc>",
            xml2
        )

        for url in urls:

            if url not in products:
                products.append(url)

    print("✅ TOTAL PRODUCTS:", len(products))

    return products



# =========================
# PRODUCT PARSER
# =========================

def parse_product(url):

    result = {
        "sku": "",
        "title": "",
        "price": "",
        "status": "",
        "url": url
    }


    try:

        r = session.get(
            url,
            timeout=60,
            headers={
                "User-Agent": "Mozilla/5.0",
                "Accept-Language": "uk-UA,uk;q=0.9",
                "Referer": BASE
            }
        )


        print(
            "STATUS PAGE:",
            r.status_code
        )


        # =====================
        # CHALLENGE
        # =====================

        if "challenge_passed" in r.text:

            print("⚠️ PRODUCT CHALLENGE")


            m = re.search(
                r'document\.cookie\s*=\s*"challenge_passed=([^"]+)',
                r.text
            )


            if m:

                session.cookies.set(
                    "challenge_passed",
                    m.group(1),
                    domain="luna-toys.com.ua",
                    path="/"
                )


                time.sleep(1)


                r = session.get(
                    url,
                    timeout=60,
                    headers={
                        "User-Agent": "Mozilla/5.0",
                        "Accept-Language": "uk-UA,uk;q=0.9",
                        "Referer": BASE
                    }
                )


        soup = BeautifulSoup(
            r.text,
            "html.parser"
        )



        # =====================
        # TITLE
        # =====================

        h1 = soup.select_one(
            "h1.product-title"
        )

        if h1:

            result["title"] = clean(
                h1.get_text()
            )



        # =====================
        # SKU
        # =====================

        sku = soup.select_one(
            ".product-header__code"
        )


        if sku:

            result["sku"] = clean(
                sku.get_text()
            ).replace(
                "Артикул:",
                ""
            ).strip()



        # =====================
        # STATUS
        # =====================

        status = soup.select_one(
            ".product-header__availability"
        )


        if status:

            result["status"] = clean(
                status.get_text()
            )



        # =====================
        # PRICE USD
        # =====================

        price = soup.select_one(
            "[itemprop='price']"
        )


        if price:

            result["price"] = price.get(
                "content",
                ""
            )


        else:

            # запасной поиск
            meta_price = soup.select_one(
                "meta[property='product:price:amount']"
            )

            if meta_price:

                result["price"] = meta_price.get(
                    "content",
                    ""
                )



        print(
            "FOUND:",
            result
        )


    except Exception as e:

        print(
            "ERROR PRODUCT:",
            url,
            e
        )


    return [
        result["sku"],
        result["title"],
        result["price"],
        result["status"],
        result["url"]
    ]
  
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


        wb = Workbook()

        ws = wb.active

        ws.append([
            "SKU",
            "TITLE",
            "PRICE",
            "STATUS",
            "URL"
        ])



        products = get_products()


        # =========================
        # TEST LIMIT
        # =========================
        
        if TEST_MODE:
            products = products[:TEST_LIMIT]
            print(f"🧪 TEST MODE: {len(products)} товаров")


        total = len(products)


        for i, url in enumerate(products,1):


            save_status(
                True,
                int(i / total * 100),
                USER,
                FILE_PATH
            )


            sku, title, price, status, link = parse_product(url)



            if not title:
                continue


            ws.append([
                sku,
                title,
                price,
                status,
                link
            ])


            time.sleep(2)



        os.makedirs(
            OUTPUT_DIR,
            exist_ok=True
        )


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


        print(
            "✅ Готово LunaBag"
        )



    finally:

        set_lock(False)


if __name__ == "__main__":
    run_parser()
