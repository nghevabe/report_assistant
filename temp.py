import pandas as pd
from openpyxl import load_workbook

from demo import exchange_by_cur
from text_util import normalize_excel_text

df = pd.read_excel("data_raw.xlsx", header=0)
df = df.filter(regex="^(?!Unnamed)")

list_keys = df.columns.tolist()

print("Quy Doi: ")
print(exchange_by_cur("EUR"))
# for item in list_keys:
#     print(normalize_excel_text(item))



