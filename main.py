import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string
from openpyxl.utils.cell import coordinate_from_string
from openpyxl.workbook import Workbook

from excel_writer import copy_excel_columns
from revenue import get_revenue_calculate_matrix, get_lst_sum_revenue_exchange

from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from copy import copy

from step_2 import copy_from_AC_to_end_and_append, find_last_data_column_any_row


def get_lst_from_col(excel_file, col_name , sheet_name=None):
    wb = load_workbook(excel_file)
    ws = wb[sheet_name] if sheet_name else wb.active

    col_index = column_index_from_string(col_name)  # cột B
    start_row = 3                             # bắt đầu từ B2

    values = []

    for row in range(start_row, ws.max_row + 1):
        values.append(ws.cell(row=row, column=col_index).value)

    return values




telegram_col = get_lst_from_col("data_raw.xlsx", "B")
date_col = get_lst_from_col("data_raw.xlsx", "C")

lst_ma_cn = []
lst_ma_pgd = []
lst_month = []

for col in date_col:
    if col is None:
        continue
    col = str(col)
    lst_month.append(col[3:5])

for col in telegram_col:
    if col is None:
        continue
    col = str(col)

    if "DN" in col:
        lst_ma_pgd.append(col[11:17])
        lst_ma_cn.append(col[11:14])
    else:
        lst_ma_pgd.append(col[8:14])
        lst_ma_cn.append(col[8:11])


from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy


# def insert_pgd_column(
#     excel_file,
#     output_file,
#     lst_ma_pgd,
#     lst_ma_cn,
#     lst_month,
#     sheet_name=None,
#     header_row=2,
#     start_data_row=3,
# ):
#     wb = load_workbook(excel_file)
#     ws = wb[sheet_name] if sheet_name else wb.active
#
#     # 1️⃣ Insert cột mới
#     ws.insert_cols(3, 1)
#
#     # 2️⃣ Ghi header
#     ws.cell(row=header_row, column=3).value = "Phòng giao dịch"
#
#     # 3️⃣ Ghi data
#     for i, value in enumerate(lst_ma_pgd):
#         cell = ws.cell(row=start_data_row + i, column=3)
#         if not isinstance(cell, MergedCell):
#             cell.value = value
#
# # ==========
#
#     # 1️⃣ Insert cột mới
#     ws.insert_cols(4, 1)
#
#     # 2️⃣ Ghi header
#     ws.cell(row=header_row, column=4).value = "Mã chi nhánh"
#
#     # 3️⃣ Ghi data
#     for i, value in enumerate(lst_ma_cn):
#         cell = ws.cell(row=start_data_row + i, column=4)
#         if not isinstance(cell, MergedCell):
#             cell.value = value
#
# # ==========
#
#     # 1️⃣ Insert cột mới
#     ws.insert_cols(5, 1)
#
#     # 2️⃣ Ghi header
#     ws.cell(row=header_row, column=5).value = "Tháng"
#
#     # 3️⃣ Ghi data
#     for i, value in enumerate(lst_month):
#         cell = ws.cell(row=start_data_row + i, column=5)
#         if not isinstance(cell, MergedCell):
#             cell.value = value
#
#     wb.save(output_file)
#     print("✅ Insert PGD column thành công")


from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy


def copy_ABC_next_to_C(
    input_file: str,
    output_file: str,
    sheet_name: str | None = None,
    header_row: int = 1,
    start_data_row: int = 2
):
    wb = load_workbook(input_file)
    ws = wb[sheet_name] if sheet_name else wb.active

    src_start_col = 1  # A
    src_end_col = 3    # C
    col_count = 3

    # 👉 chèn 3 cột NGAY SAU C
    insert_at_col = src_end_col + 1  # D
    ws.insert_cols(insert_at_col, amount=col_count)

    max_row = ws.max_row

    print("Insert 3 columns at D, then copy A,B,C → D,E,F")

    # ===============================
    # 1️⃣ Copy column width
    # ===============================
    for i in range(col_count):
        ws.column_dimensions[get_column_letter(insert_at_col + i)].width = \
            ws.column_dimensions[get_column_letter(src_start_col + i)].width

    # ===============================
    # 2️⃣ Copy header
    # ===============================
    for i in range(col_count):
        src_cell = ws.cell(row=header_row, column=src_start_col + i)
        dest_cell = ws.cell(row=header_row, column=insert_at_col + i)

        dest_cell.value = src_cell.value

        if src_cell.has_style:
            dest_cell.font = copy(src_cell.font)
            dest_cell.border = copy(src_cell.border)
            dest_cell.fill = copy(src_cell.fill)
            dest_cell.number_format = copy(src_cell.number_format)
            dest_cell.alignment = copy(src_cell.alignment)

    # ===============================
    # 3️⃣ Copy data
    # ===============================
    for row in range(start_data_row, max_row + 1):
        for i in range(col_count):
            src_cell = ws.cell(row=row, column=src_start_col + i)
            dest_cell = ws.cell(row=row, column=insert_at_col + i)

            if isinstance(src_cell, MergedCell) or isinstance(dest_cell, MergedCell):
                continue

            dest_cell.value = src_cell.value

            if src_cell.has_style:
                dest_cell.font = copy(src_cell.font)
                dest_cell.border = copy(src_cell.border)
                dest_cell.fill = copy(src_cell.fill)
                dest_cell.number_format = copy(src_cell.number_format)
                dest_cell.alignment = copy(src_cell.alignment)

    update_table(output_file, ws, wb, lst_ma_pgd, lst_ma_cn, lst_month)
    wb.save(output_file)
    print(f"✅ Done. Saved to {output_file}")


def update_table(output_file, ws, wb, lst_ma_pgd, lst_ma_cn, lst_month):

    row_index = 3
    for item in lst_ma_pgd:
        ws.cell(row=row_index, column=4).value = item
        row_index += 1

    row_index = 3
    for item in lst_ma_cn:
        ws.cell(row=row_index, column=5).value = item
        row_index += 1

    row_index = 3
    for item in lst_month:
        ws.cell(row=row_index, column=6).value = item
        row_index += 1

    ws["D2"].value = "Phòng giao dịch"
    ws["E2"].value = "Mã chi nhánh"
    ws["F2"].value = "Tháng"
    wb.save(output_file)
    print(f"✅ Done. Output saved to: {output_file}")


# DN_26122025120155271622
# 26122025279153271365

def copy_range_to_new_file(
    src_file: str,
    dst_file: str,
    src_sheet_name: str | None = None,
    dst_sheet_name: str = "Sheet1",
    start_col_letter: str = "A",
    end_col_letter: str = "AE",
    start_row: int = 2
):
    # ===============================
    # Load source workbook
    # ===============================
    src_wb = load_workbook(src_file)
    src_ws = src_wb[src_sheet_name] if src_sheet_name else src_wb.active

    # ===============================
    # Create destination workbook
    # ===============================
    dst_wb = Workbook()
    dst_ws = dst_wb.active
    dst_ws.title = dst_sheet_name

    start_col = column_index_from_string(start_col_letter)
    end_col = column_index_from_string(end_col_letter)

    max_row = src_ws.max_row

    print(f"Copy range {start_col_letter}{start_row}:{end_col_letter}{max_row}")

    # ===============================
    # 1️⃣ Copy column width
    # ===============================
    for col in range(start_col, end_col + 1):
        col_letter = get_column_letter(col)
        dst_ws.column_dimensions[col_letter].width = \
            src_ws.column_dimensions[col_letter].width

    # ===============================
    # 2️⃣ Copy cells (value + style)
    # ===============================
    for row in range(start_row, max_row + 1):
        for col in range(start_col, end_col + 1):
            src_cell = src_ws.cell(row=row, column=col)
            dst_cell = dst_ws.cell(
                row=row - start_row + 1,   # dán từ row 1 ở file mới
                column=col - start_col + 1
            )

            # Skip read-only merged cells
            if isinstance(src_cell, MergedCell):
                continue

            dst_cell.value = src_cell.value

            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.border = copy(src_cell.border)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.number_format = copy(src_cell.number_format)
                dst_cell.alignment = copy(src_cell.alignment)
                dst_cell.protection = copy(src_cell.protection)

    # ===============================
    # 3️⃣ Copy merged cells
    # ===============================
    merged_ranges = list(src_ws.merged_cells.ranges)

    for merged in merged_ranges:
        min_col, min_row, max_col, max_row_m = merged.bounds

        # chỉ copy merge nằm trong vùng A2:AE*
        if (
            min_col >= start_col and
            max_col <= end_col and
            max_row_m >= start_row
        ):
            dst_ws.merge_cells(
                start_row=min_row - start_row + 1,
                start_column=min_col - start_col + 1,
                end_row=max_row_m - start_row + 1,
                end_column=max_col - start_col + 1,
            )

    # ===============================
    # Save destination file
    # ===============================
    dst_wb.save(dst_file)
    print(f"✅ Done. Saved to {dst_file}")


def insert_sum_column(
    excel_file,
    output_file,
    list_revenue,
    sheet_name=None,
    header_row=2,
):
    wb = load_workbook(excel_file)
    ws = wb[sheet_name] if sheet_name else wb.active

    col_number = find_last_data_column_any_row(ws)

    # 1️⃣ Insert cột mới
    ws.insert_cols(col_number+5, 1)

    # 2️⃣ Ghi header
    ws.cell(row=header_row, column=col_number+5).value = "Tổng doanh số"

    # 3️⃣ Ghi data
    for i, value in enumerate(list_revenue):
        cell = ws.cell(row=3 + i, column=col_number+5)
        if not isinstance(cell, MergedCell):
            cell.value = value

    wb.save(output_file)
    print("✅ Tính tổng doanh số thành công thành công")


lst_revenue = get_lst_sum_revenue_exchange()

insert_sum_column(
    excel_file="result_output_final.xlsx",
    output_file="result_output_final.xlsx",
    list_revenue=lst_revenue,
)



# copy_ABC_next_to_C(
#         input_file="data_raw.xlsx",
#         output_file="result_step_1.xlsx"
#     )
#
# copy_range_to_new_file(
#         src_file="result_step_1.xlsx",
#         dst_file="result_output_final.xlsx",
#         start_col_letter="A",
#         end_col_letter="AE",
#         start_row=2
#     )
#
# copy_excel_columns(
#     input_file="data_raw.xlsx",
#     output_file="result_output_2.xlsx",
#     sheet_name=None,
#     src_start_col_letter="AC",
#     header_row=1,
#     gap_cols=5,
#     matrix_revenue=get_revenue_calculate_matrix(lst_ma_pgd)
# )
#
#
# copy_from_AC_to_end_and_append(
#         src_file="result_output_2.xlsx",
#         dst_file="result_output_final.xlsx",
#         start_col_letter="AC",
#         header_row=1,
#         start_data_row=2
#     )

