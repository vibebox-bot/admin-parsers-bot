import os
from openpyxl import load_workbook, Workbook

# =========================
# ФАЙЛЫ
# =========================

CRM_FILE = "checker/export-2026-07-12_00-14-43.xlsx"
MELAD_FILE = "output/Melad/Melad_LIVE.xlsx"
RESULT_FILE = "output/Melad/Melad_CHECK.xlsx"


# =========================
# CRM EXCEL
# =========================

def get_salesdrive():

    wb = load_workbook(CRM_FILE, data_only=True)
    ws = wb.active

    headers = {}

    for i, cell in enumerate(ws[1], start=1):
        headers[str(cell.value).strip()] = i

    print("📥 SalesDrive Excel")
    print(headers)

    supplier_col = headers["Постачальник"]
    barcode_col = headers["Штрихкод"]
    note_col = headers["Нотатка"]
    cost_col = headers["Собівартість"]
    price_col = headers["Ціна"]

    products = {}

    for row in ws.iter_rows(min_row=2, values_only=True):

        supplier = row[supplier_col - 1]

        if str(supplier).strip() != "Melad":
            continue

        article = ""

        barcode = row[barcode_col - 1]
        note = row[note_col - 1]

        if barcode not in ("", None):
            article = str(barcode).strip()

        elif note not in ("", None):
            article = str(note).strip()

        if article == "":
            continue

        cost = row[cost_col - 1]

        try:
            site_price = float(row[price_col - 1])
        except:
            site_price = 0

        if site_price <= 1:
            stock = 0
        else:
            stock = 1

        products[article] = {
            "cost": cost,
            "stock": stock,
            "site_price": site_price
        }

    print("Melad в CRM:", len(products))

    return products


# =========================
# MELAD PARSER
# =========================

def get_melad():

    wb = load_workbook(MELAD_FILE, data_only=True)
    ws = wb.active

    products = {}

    for row in ws.iter_rows(min_row=2, values_only=True):

        name = row[0]
        price = row[1]
        article = row[2]
        stock = row[3]
        url = row[4]

        if not article:
            continue

        article = str(article).strip()

        # Цена
        price = str(price).replace("$", "")
        price = price.replace("грн", "")
        price = price.replace(",", ".")
        price = price.strip()

        try:
            price = float(price)
        except:
            price = 0

        # Наличие
        try:
            stock = int(float(stock))
        except:
            stock = 0

        stock = 1 if stock > 0 else 0

        products[article] = {
            "name": name,
            "price": price,
            "stock": stock,
            "url": url
        }

    print("📥 Melad parser")
    print("Melad товаров:", len(products))

    return products

# =========================
# MAIN
# =========================

def main():

    crm = get_salesdrive()
    melad = get_melad()

    wb = Workbook()
    ws = wb.active
    ws.title = "Melad"

    ws.append([
        "Артикул",
        "Название",
        "Цена парсер ($)",
        "Себестоимость CRM ($)",
        "Наличие парсер",
        "Наличие CRM",
        "Статус",
        "URL"
    ])

    changed_price = 0
    changed_stock = 0
    ok = 0
    missing = 0

    for article, item in melad.items():

        crm_item = crm.get(article)

        if crm_item is None:

            ws.append([
                article,
                item["name"],
                item["price"],
                "",
                item["stock"],
                "",
                "❌ Нет в CRM",
                item["url"]
            ])

            missing += 1
            continue

        # ---------- Цена ----------

        parser_price = item["price"]

        try:
            crm_cost = float(str(crm_item["cost"]).replace(",", "."))
        except:
            crm_cost = 0

        # ---------- Наличие ----------

        parser_stock = 0 if str(item["stock"]).strip() == "0" else 1
        crm_stock = crm_item["stock"]

        status = []

        if abs(parser_price - crm_cost) > 0.001:
            status.append("💲 Цена")

        if parser_stock != crm_stock:
            status.append("📦 Наличие")

        if not status:
            status.append("✅ OK")
            ok += 1
        else:
            if "💲 Цена" in status:
                changed_price += 1

            if "📦 Наличие" in status:
                changed_stock += 1

        ws.append([
            article,
            item["name"],
            parser_price,
            crm_cost,
            parser_stock,
            crm_stock,
            ", ".join(status),
            item["url"]
        ])

    wb.save(RESULT_FILE)

    print()
    print("==========")
    print("OK:", ok)
    print("Цена:", changed_price)
    print("Наличие:", changed_stock)
    print("Нет в CRM:", missing)
    print("==========")
    print(RESULT_FILE)

if __name__ == "__main__":
    main()
