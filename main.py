import pandas as pd
from openpyxl import load_workbook

from revenue import revenue_calculate

df = pd.read_excel("data_raw.xlsx", header=1)

# Chuẩn hóa tên cột
df.columns = df.columns.astype(str).str.strip()

telegram_col = df["Số điện"].dropna().astype(str)
date_col = df["Ngày hiệu lực"].dropna().astype(str)

lst_ma_cn = []
lst_ma_pgd = []
lst_month = []

for col in date_col:
    lst_month.append(col[3:5])

for col in telegram_col:
    if "DN" in col:
        lst_ma_pgd.append(str(col[11:17]))
        lst_ma_cn.append(str(col[11:14]))
    else:
        lst_ma_pgd.append(str(col[8:14]))
        lst_ma_cn.append(str(col[8:11]))

# print(lst_ma_pgd)

revenue_calculate(lst_ma_pgd)



# wb = load_workbook("data_raw.xlsx")
# ws = wb.active
#
# HEADER_ROW = 2
#
# INSERT_COL_PGD = 3   # cột thứ 3 (C)
#
# ws.insert_cols(INSERT_COL_PGD)
# ws.cell(row=HEADER_ROW, column=INSERT_COL_PGD).value = "Phòng giao dịch"
#
# start_data_row = HEADER_ROW + 1
#
# for i, value in enumerate(lst_ma_pgd):
#     ws.cell(
#         row=start_data_row + i,
#         column=INSERT_COL_PGD
#     ).value = value
#
# # =========
#
# INSERT_COL_CN = 4   # cột thứ 4 (C)
#
# ws.insert_cols(INSERT_COL_CN)
# ws.cell(row=HEADER_ROW, column=INSERT_COL_CN).value = "Mã chi nhánh"
#
# start_data_row = HEADER_ROW + 1
#
# for i, value in enumerate(lst_ma_cn):
#     ws.cell(
#         row=start_data_row + i,
#         column=INSERT_COL_CN
#     ).value = value
#
# # =========
#
# INSERT_COL_CN = 6   # cột thứ 4 (C)
#
# ws.insert_cols(INSERT_COL_CN)
# ws.cell(row=HEADER_ROW, column=INSERT_COL_CN).value = "Tháng"
#
# start_data_row = HEADER_ROW + 1
#
# for i, value in enumerate(lst_month):
#     ws.cell(
#         row=start_data_row + i,
#         column=INSERT_COL_CN
#     ).value = value
#
# wb.save("data_output.xlsx")

# DN_26122025120155271622
# 26122025279153271365



