import os
import requests
import xml.etree.ElementTree as ET
from openpyxl import load_workbook, Workbook


YML_URL = "https://kindlytech.salesdrive.me/export/yml/export.yml?publicKey=YEWxvIKV_z6Hjx4-zqWiGLmmsFAS05TLQQ23qZbeoR_2UjOCNEtx-QxFfP0JFfUv45Q"

MELAD_FILE = "output/Melad/Melad_LIVE.xlsx"

RESULT_FILE = "output/Melad/Melad_CHECK.xlsx"


# =========================
# SALES DRIVE
# =========================

def get_salesdrive():

    r = requests.get(
        YML_URL,
        timeout=60
    )

    r.raise_for_status()

    root = ET.fromstring(r.text)

    products = {}

    for offer in root.findall(".//offer"):

        article = offer.findtext(
            "article",
            ""
        ).strip()

        if not article:
            continue

        keys = []
        
        if offer.findtext("note"):
            keys.append(offer.findtext("note").strip())
        
        if offer.findtext("barcode"):
            keys.append(offer.findtext("barcode").strip())
        
        
        for key in keys:
        
            products[key] = {
        
                "name": offer.findtext("name",""),
        
                "stock": offer.findtext(
                    "quantity_in_stock",
                    "0"
                ),
        
                "price": offer.findtext(
                    "price",
                    ""
                ),
        
                "note": offer.findtext(
                    "note",
                    ""
                )
            }

            "name": offer.findtext("name",""),

            "stock": offer.findtext(
                "quantity_in_stock",
                "0"
            ),

            "price": offer.findtext(
                "price",
                ""
            ),

            "note": offer.findtext(
                "note",
                ""
            )
        }


    return products



# =========================
# MELAD
# =========================

def get_melad():

    wb = load_workbook(
        MELAD_FILE
    )

    ws = wb.active


    products = {}

    for row in ws.iter_rows(
        min_row=2,
        values_only=True
    ):

        name, price, article, stock, url = row


        if article:

            products[str(article)] = {

                "name": name,

                "price": price,

                "stock": stock,

                "url": url
            }


    return products



# =========================
# COMPARE
# =========================

def main():

    print("📥 SalesDrive...")
    crm = get_salesdrive()

    print(
        "CRM товаров:",
        len(crm)
    )


    print("📥 Melad...")
    melad = get_melad()

    print(
        "Melad товаров:",
        len(melad)
    )


    wb = Workbook()

    ws = wb.active

    ws.title = "Проверка"


    ws.append([
        "Название",
        "Артикул",
        "Цена Melad",
        "Наличие Melad",
        "Наличие CRM",
        "Цена CRM",
        "Статус"
    ])


    count = 0


    for article, item in melad.items():

        crm_item = crm.get(article)


        if crm_item:

            status = "OK"

            if str(crm_item["stock"]) == "0":
                status = "НЕТ В CRM"


            ws.append([

                item["name"],

                article,

                item["price"],

                item["stock"],

                crm_item["stock"],

                crm_item["price"],

                status
            ])

            count += 1



    wb.save(
        RESULT_FILE
    )


    print(
        "ГОТОВО:",
        count
    )

    print(
        RESULT_FILE
    )



if __name__ == "__main__":
    main()
