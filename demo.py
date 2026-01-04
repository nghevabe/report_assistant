import requests

API_KEY = "fa384bbcf530039925efa6d5"
url = f"https://v6.exchangerate-api.com/v6/{API_KEY}/latest/USD"


def exchange_by_cur(currency_code):
    res = requests.get(url, timeout=5)
    data = res.json()

    if data.get("result") != "success":
        raise RuntimeError(data)

    rates = data["conversion_rates"]
    print(rates[currency_code])


