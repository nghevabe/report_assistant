import pandas as pd

from text_util import normalize_excel_text

df_raw = pd.read_excel("data_raw.xlsx", header=None)

KEY_HEADER_CON = "Số điện"

header_con_row = None

for r in range(len(df_raw)):
    if KEY_HEADER_CON in df_raw.iloc[r].astype(str).values:
        header_con_row = r
        break

if header_con_row is None:
    raise ValueError("Không tìm thấy Header con")

header_cha_row = header_con_row - 1 if header_con_row > 0 else None

header_con = df_raw.iloc[header_con_row]
header_cha = (
    df_raw.iloc[header_cha_row]
    if header_cha_row is not None
    else [None] * len(header_con)
)

final_headers = []

for cha, con in zip(header_cha, header_con):
    cha = str(cha).strip() if pd.notna(cha) else ""
    con = str(con).strip() if pd.notna(con) else ""

    if cha and con:
        final_headers.append(f"{normalize_excel_text(cha)} - {normalize_excel_text(con)}")
    elif con:
        final_headers.append(con)
    elif cha:
        final_headers.append(cha)
    else:
        final_headers.append("")

print(final_headers)