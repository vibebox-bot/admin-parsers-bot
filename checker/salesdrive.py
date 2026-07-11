import requests

API_KEY = "iCqBJpalnO_rAtMXmIcaGdUelJP7HFOOm41kxD0FVxH9fo0jyBCcGyQNF12YYy3ppdDRmMsQjxn9fXyW4Dkyz4YJYArSEP_ADUdF"

url = "https://kindlytech.salesdrive.me/api/products"

headers = {
    "X-Api-Key": API_KEY
}

r = requests.get(
    url,
    headers=headers
)

print(r.status_code)
print(r.text[:500])
