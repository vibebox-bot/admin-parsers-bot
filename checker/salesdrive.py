import requests
import xml.etree.ElementTree as ET


YML_URL = "https://kindlytech.salesdrive.me/export/yml/export.yml?publicKey=YEWxvIKV_z6Hjx4-zqWiGLmmsFAS05TLQQ23qZbeoR_2UjOCNEtx-QxFfP0JFfUv45Q"


def get_products():

    r = requests.get(YML_URL, timeout=60)

    r.raise_for_status()

    root = ET.fromstring(r.text)

    products = []

    for offer in root.findall(".//offer"):

        product = {

            "name": offer.findtext("name", ""),

            "article": offer.findtext("article", ""),

            "barcode": offer.findtext("barcode", ""),

            "note": offer.findtext("note", ""),

            "vendorprice": offer.findtext("vendorprice", ""),

            "stock": offer.findtext("quantity_in_stock", ""),

            "price": offer.findtext("price", ""),

        }

        products.append(product)


    return products

if __name__ == "__main__":

    items = get_products()

    print("ТОВАРОВ:", len(items))

    print(items[0])
