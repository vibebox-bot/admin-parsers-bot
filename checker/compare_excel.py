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
    print("Колонки найдены:")
    print(headers)

    supplier_col = headers["Постачальник"]
    sku_col = headers["SKU"]
    cost_col = headers["Собівартість"]
    stock_col = headers["Залишок на складі"]

    products = {}

    for row in ws.iter_rows(min_row=2, values_only=True):

        supplier = row[supplier_col - 1]

        if supplier != "Melad":
            continue

        sku = row[sku_col - 1]

        if not sku:
            continue

        products[str(sku).strip()] = {
            "cost": row[cost_col - 1],
            "stock": row[stock_col - 1]
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

        products[str(article).strip()] = {
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
        "Цена поставщика",
        "Себестоимость CRM",
        "Наличие поставщика",
        "Наличие CRM",
        "URL",
        "Статус"
    ])

    ok = 0
    missing = 0

    for article, item in melad.items():

        crm_item = crm.get(article)

        if crm_item:

            supplier_stock = str(item["stock"]).strip()

            try:
                crm_stock = int(float(crm_item["stock"]))
            except:
                crm_stock = 0

            if supplier_stock in ("", "0", "Нет", "нет", "Немає"):
                supplier_have = 0
            else:
                supplier_have = 1

            if supplier_have == crm_stock:
                status = "OK"
            elif supplier_have == 1 and crm_stock == 0:
                status = "❌ НЕТ В CRM"
            elif supplier_have == 0 and crm_stock == 1:
                status = "➕ ПОЯВИЛСЯ"
            else:
                status = "⚠ ПРОВЕРИТЬ"

            ws.append([
                article,
                item["name"],
                item["price"],
                crm_item["cost"],
                supplier_have,
                crm_stock,
                item["url"],
                status
            ])

            ok += 1

        else:

            ws.append([
                article,
                item["name"],
                item["price"],
                "",
                item["stock"],
                "",
                item["url"],
                "❌ НЕТ В CRM"
            ])

            missing += 1

    os.makedirs("output/Melad", exist_ok=True)
    wb.save(RESULT_FILE)

    print()
    print("ГОТОВО")
    print("Совпало:", ok)
    print("Нет в CRM:", missing)
    print(RESULT_FILE)


if __name__ == "__main__":
    main()
