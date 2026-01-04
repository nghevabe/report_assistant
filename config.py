import requests

API_KEY = "fa384bbcf530039925efa6d5"
url_api = f"https://v6.exchangerate-api.com/v6/{API_KEY}/latest/USD"


def get_exchange(url):
    res = requests.get(url, timeout=5)
    return res.json()


def exchange_by_cur(data, currency_code):
    if data.get("result") != "success":
        raise RuntimeError(data)

    rates = data["conversion_rates"]
    return rates[currency_code]
