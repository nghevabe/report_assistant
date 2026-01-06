import re
import pandas as pd


def normalize_excel_text(value: str) -> str:
    if value is None:
        return ""

    s = str(value)

    # Remove XML null escape
    s = re.sub(r"_x0000_", "", s)

    # Remove actual null characters
    s = s.replace("\x00", "")

    # Normalize spaces
    s = s.strip()

    return s
