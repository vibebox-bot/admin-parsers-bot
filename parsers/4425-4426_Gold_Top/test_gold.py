import requests

url = "https://gold-tor.com.ua/"

headers = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/131.0.0.0 Safari/537.36"
    ),
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.7",
}

try:

    r = requests.get(
        url,
        headers=headers,
        timeout=30,
        allow_redirects=True
    )

    print("========== GOLD TEST ==========")
    print("STATUS:", r.status_code)
    print("URL:", r.url)
    print("SERVER:", r.headers.get("Server"))
    print("RETRY-AFTER:", r.headers.get("Retry-After"))
    print("CONTENT-TYPE:", r.headers.get("Content-Type"))
    print("CONTENT-LENGTH:", r.headers.get("Content-Length"))
    print("DATE:", r.headers.get("Date"))
    print("VIA:", r.headers.get("Via"))
    print("CF-RAY:", r.headers.get("CF-RAY"))
    print("X-POWERED-BY:", r.headers.get("X-Powered-By"))
    print("===============================")

except Exception as e:

    print("TEST ERROR:", repr(e))
