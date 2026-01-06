import re

import pandas as pd

from config import get_exchange, url_api, exchange_by_cur
from text_util import normalize_excel_text

df = pd.read_excel("data_raw.xlsx", header=1)
df = df.filter(regex="^(?!Unnamed)")

list_keys_all = df.columns.tolist()
lst_key_money = list_keys_all[28:]
lst_key_money_str = [str(x) for x in lst_key_money]


face_value_data = pd.read_excel("data_raw.xlsx", header=0)
face_value_data = face_value_data.filter(regex="^(?!Unnamed)")

list_face_value_raw = face_value_data.columns.tolist()
list_face_value = []

for item in list_face_value_raw:
    normalize_excel_text(item)
    list_face_value.append(normalize_excel_text(item))

lst_money_value = [re.sub(r"\.\d+$", "", x).replace(",", "") for x in lst_key_money_str]
lst_money_value.pop()

lst_flat_face_money = []

face_money_index = 0
for value in lst_money_value:
    if "Thành tiền" in value:
        face_money_index = face_money_index+1
    result = list_face_value[face_money_index] + "#" + value
    lst_flat_face_money.append(result)

lst_flat_face_money = lst_flat_face_money[:15]


def lst_banknotes_by_pgd(lst_ma_pgd, banknote):
    lst_revenue = []
    for ma_pgd in lst_ma_pgd:
        revenue = 0
        all_indices = [index for index, item in enumerate(lst_ma_pgd) if item == ma_pgd]
        for index in all_indices:
            revenue = revenue + banknote[index]
        lst_revenue.append(ma_pgd + "-" + str(revenue))

    return set(lst_revenue)


def revenue_calculate(lst_pgd):
    data_exchange = get_exchange(url_api)

    face_value_index = 0
    for face_money_flat in lst_flat_face_money:
        face_value_index += 1
        if "Thành tiền" not in face_money_flat:
            lst_banknotes = df.iloc[:, 27 + face_value_index].tolist()
            lst_banknotes = lst_banknotes[:len(lst_banknotes)-1]

            pgd_index = 0
            for banknote_item in lst_banknotes:
                currency_code = face_money_flat.split("#")[0]
                face_money_value = face_money_flat.split("#")[1]
                exchange_value = exchange_by_cur(data_exchange, currency_code)

                revenue = int(banknote_item) * int(face_money_value) / exchange_value
                revenue = round(revenue, 2)
                print(str(lst_pgd[pgd_index])+"#"+str(banknote_item)+"#"+str(currency_code)+"#"+str(face_money_value)+"="+str(revenue))
                pgd_index += 1