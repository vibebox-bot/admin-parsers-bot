import requests

url = "https://gold-tor.com.ua/"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": (
        "text/html,application/xhtml+xml,application/xml;"
        "q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en-US;q=0.7,en;q=0.6",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
}

try:

    r = requests.get(
        url,
        headers=headers,
        timeout=30,
        allow_redirects=True
    )

    print("==============================")
    print("STATUS:", r.status_code)
    print("URL:", r.url)
    print("SERVER:", r.headers.get("Server"))
    print("RETRY-AFTER:", r.headers.get("Retry-After"))
    print("CONTENT-TYPE:", r.headers.get("Content-Type"))
    print("CONTENT-LENGTH:", r.headers.get("Content-Length"))
    print("==============================")
    print("TEXT:")
    print(r.text[:2000])
    print("==============================")

except Exception as e:

    print("ERROR:", repr(e))
