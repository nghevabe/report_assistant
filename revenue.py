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
    result = list_face_value[face_money_index] + "#" + value
    lst_flat_face_money.append(result)
    if "Thành tiền" in value:
        face_money_index = face_money_index + 1


# lst_flat_face_money = lst_flat_face_money[:15]


def lst_banknotes_by_pgd(lst_ma_pgd, banknote):
    lst_revenue = []
    for ma_pgd in lst_ma_pgd:
        revenue = 0
        all_indices = [index for index, item in enumerate(lst_ma_pgd) if item == ma_pgd]
        for index in all_indices:
            revenue = revenue + banknote[index]
        lst_revenue.append(ma_pgd + "-" + str(revenue))

    return set(lst_revenue)


def get_revenue_calculate_matrix(lst_pgd):
    # data_exchange = get_exchange(url_api)
    matrix_revenue_face_value = []

    face_value_index = 0
    for face_money_flat in lst_flat_face_money:
        lst_revenue_face_value = []
        face_value_index += 1
        if "Thành tiền" not in face_money_flat:
            lst_banknotes = df.iloc[:, 27 + face_value_index].tolist()
            lst_banknotes = lst_banknotes[:len(lst_banknotes) - 1]

            pgd_index = 0
            for banknote_item in lst_banknotes:
                currency_code = face_money_flat.split("#")[0]
                face_money_value = face_money_flat.split("#")[1]
                # exchange_value = exchange_by_cur(data_exchange, currency_code)

                revenue = int(banknote_item) * int(face_money_value)
                revenue = round(revenue, 2)
                result_str = str(lst_pgd[pgd_index]) + "#" + str(banknote_item) + "#" + str(currency_code) + "#" + str(
                    face_money_value) + "=" + str(revenue)
                # print(result_str)
                pgd_index += 1
                lst_revenue_face_value.append(result_str)

        matrix_revenue_face_value.append(lst_revenue_face_value)

    print(matrix_revenue_face_value)
    return matrix_revenue_face_value


def get_lst_sum_revenue_exchange():
    data_exchange = get_exchange(url_api)
    matrix_sum_revenue_by_exchange = []

    face_value_index = 0
    for face_money_flat in lst_flat_face_money:
        is_dont_need = ""
        lst_sum_revenue_exchange_by_face_value = []
        face_value_index += 1
        if "Thành tiền" in face_money_flat:
            currency_code_str = face_money_flat.split("#")[0]
            if "-" in currency_code_str:
                is_dont_need = "x"
                currency_code = currency_code_str.split("-")[0]
            else:
                currency_code = currency_code_str
            exchange_value = exchange_by_cur(data_exchange, currency_code)
            lst_revenues = df.iloc[:, 27 + face_value_index].tolist()
            lst_revenues = lst_revenues[:len(lst_revenues) - 1]

            for revenue in lst_revenues:
                revenue_exchange = int(revenue) / exchange_value
                revenue_exchange = round(revenue_exchange, 2)
                lst_sum_revenue_exchange_by_face_value.append(str(revenue_exchange) + is_dont_need)

            # print(face_money_flat)
            # print(exchange_value)
            # print(lst_sum_revenue_exchange)
            matrix_sum_revenue_by_exchange.append(lst_sum_revenue_exchange_by_face_value)

    lst_sum_revenue_exchange = []
    for x in range(len(matrix_sum_revenue_by_exchange[0])):
        # print("Col:")
        sum_revenue = 0
        for y in range(len(matrix_sum_revenue_by_exchange)):
            if "x" not in matrix_sum_revenue_by_exchange[y][x]:
                sum_revenue += float(matrix_sum_revenue_by_exchange[y][x])
        lst_sum_revenue_exchange.append(sum_revenue)

    return lst_sum_revenue_exchange
