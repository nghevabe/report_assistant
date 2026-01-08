from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.cell.cell import MergedCell
from copy import copy


def find_last_data_column_any_row(ws):
    """
    Tìm cột cuối cùng có data ở BẤT KỲ hàng nào
    (bỏ qua header trống / merge)
    """
    max_col = ws.max_column
    max_row = ws.max_row

    for col in range(max_col, 0, -1):
        for row in range(1, max_row + 1):
            if ws.cell(row=row, column=col).value not in (None, ""):
                return col
    return 0



def safe_append_start_col(ws):
    """Chọn cột bắt đầu paste sao cho KHÔNG nằm trong merge sẵn có ở sheet đích."""
    max_end = ws.max_column
    for mr in ws.merged_cells.ranges:
        _, _, end_col, _ = mr.bounds
        if end_col > max_end:
            max_end = end_col
    return max_end + 1


def compute_min_row_needed_for_merges(ws, start_col, end_col, default_start_row):
    """Đảm bảo copy luôn cả top-left của các merge thuộc vùng cột cần copy."""
    min_row = default_start_row
    for mr in ws.merged_cells.ranges:
        min_c, min_r, max_c, _ = mr.bounds
        if min_c >= start_col and max_c <= end_col:
            if min_r < min_row:
                min_row = min_r
    return min_row


def copy_from_AC_to_end_and_append(
    src_file: str,
    dst_file: str,
    src_sheet_name: str | None = None,
    dst_sheet_name: str | None = None,
    start_col_letter: str = "AC",
    header_row: int = 1,
    start_data_row: int = 2,
):
    # Load source
    src_wb = load_workbook(src_file)
    src_ws = src_wb[src_sheet_name] if src_sheet_name else src_wb.active

    src_start_col = column_index_from_string(start_col_letter)
    src_end_col = find_last_data_column_any_row(src_ws)

    if src_end_col < src_start_col:
        raise ValueError("Không có dữ liệu từ cột AC trở đi (dựa theo header_row).")

    # ✅ quan trọng: copy từ hàng nhỏ nhất cần thiết để không mất top-left của merge
    copy_start_row = min(header_row, start_data_row)
    copy_start_row = compute_min_row_needed_for_merges(
        src_ws, src_start_col, src_end_col, copy_start_row
    )

    max_row = src_ws.max_row
    col_count = src_end_col - src_start_col + 1

    print(
        f"Source range: {start_col_letter}{copy_start_row}:"
        f"{get_column_letter(src_end_col)}{max_row}"
    )

    # Load destination
    dst_wb = load_workbook(dst_file)
    dst_ws = dst_wb[dst_sheet_name] if dst_sheet_name else dst_wb.active

    # ✅ quan trọng: tìm vị trí append an toàn sau mọi merge hiện có
    dest_start_col = safe_append_start_col(dst_ws)

    print(f"Paste start at: {get_column_letter(dest_start_col)}{copy_start_row}")

    # 1) Copy column width
    for i in range(col_count):
        s_col = src_start_col + i
        d_col = dest_start_col + i
        dst_ws.column_dimensions[get_column_letter(d_col)].width = \
            src_ws.column_dimensions[get_column_letter(s_col)].width

    # 2) Copy value + style (skip MergedCell read-only)
    for row in range(copy_start_row, max_row + 1):
        for i in range(col_count):
            src_cell = src_ws.cell(row=row, column=src_start_col + i)
            dst_cell = dst_ws.cell(row=row, column=dest_start_col + i)

            if isinstance(src_cell, MergedCell) or isinstance(dst_cell, MergedCell):
                continue

            dst_cell.value = src_cell.value

            if src_cell.has_style:
                dst_cell.font = copy(src_cell.font)
                dst_cell.border = copy(src_cell.border)
                dst_cell.fill = copy(src_cell.fill)
                dst_cell.number_format = copy(src_cell.number_format)
                dst_cell.alignment = copy(src_cell.alignment)
                dst_cell.protection = copy(src_cell.protection)

    # 3) Copy merged ranges (nguyên khối)
    for mr in list(src_ws.merged_cells.ranges):
        min_c, min_r, max_c, max_r = mr.bounds
        if min_c >= src_start_col and max_c <= src_end_col and max_r >= copy_start_row:
            dst_ws.merge_cells(
                start_row=min_r,
                end_row=max_r,
                start_column=dest_start_col + (min_c - src_start_col),
                end_column=dest_start_col + (max_c - src_start_col),
            )

    dst_wb.save(dst_file)
    print(f"✅ Done. Appended to {dst_file}")
