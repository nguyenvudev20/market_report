#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
replace_fields.py
- Thay thế giá trị cho từng cột theo mapping (case-insensitive).
- Thay thế có điều kiện (ví dụ: nếu Mode bắt đầu bằng 'HI' thì Manufacturer='Hanna').
- Áp dụng cho tất cả sheet trong workbook. Chỉ chạy ở sheet có cột tương ứng.

Cách chạy:
    pip install pandas xlsxwriter
    python replace_fields.py
"""

import pandas as pd
import numpy as np
import unicodedata
import re
from collections import Counter
from pathlib import Path

# ---- CẤU HÌNH ----
INPUT_XLSX  = "Vietnam_2-9_2025_Market Share Report.xlsx"
OUTPUT_XLSX = "Vietnam_2-9_2025_Market Share Report_REPLACED.xlsx"

# 1) Mapping cho từng cột (ví dụ & gợi ý – bạn sửa/ bổ sung theo dữ liệu)
MAPPINGS_MANUFACTURER = {
    # bên trái: các biến thể cần thay (không phân biệt hoa/thường)
    # bên phải: giá trị đích
    "hach co.": "Hach",
    "hach": "Hach",
    "hanna instruments": "Hanna",
    "hanna": "Hanna",
    "thermo fisher": "Thermo Fisher",
    "thermo fisher scientific": "Thermo Fisher",
    # ...
}
MAPPINGS_INSTRUMENT_TYPE = {
    "bench top": "Benchtop",
    "benchtop": "Benchtop",
    "portable": "Portable",
    "handheld": "Handheld",
    # ...
}
MAPPINGS_PARAMETER_METHOD = {
    "ph": "pH",
    "uv vis": "UV-Vis",
    "uv-vis": "UV-Vis",
    "hplc": "HPLC",
    "gc-ms": "GC-MS",
    # ...
}

# 2) Các rule có điều kiện (mỗi rule là 1 dict)
#   op hỗ trợ: equals | contains | startswith | endswith | regex | inlist
RULES = [
    {
        "if_col": "Mode",
        "op": "startswith",
        "value": "HI",
        "set_col": "Manufacturer",
        "set_value": "Hanna",
        "case_insensitive": True,
        "audit": True,  # thêm cột <set_col>__old + cờ __rule_xx
    },
    # Ví dụ khác:
    # {"if_col": "Instrument Type", "op": "equals", "value": "Benchtop",
    #  "set_col": "Parameter/Method", "set_value": "UV-Vis", "case_insensitive": True, "audit": True},
]

# ---- HÀM TIỆN ÍCH ----
def norm_text(x):
    """Unicode NFKC + strip; trả về '' nếu NaN/None."""
    if x is None or (isinstance(x, float) and np.isnan(x)):
        return ""
    s = str(x)
    s = unicodedata.normalize("NFKC", s)
    return s.strip()

def find_col(df, want_name: str):
    """Tìm cột theo tên không phân biệt hoa/thường; trả về tên cột thật trong df hoặc ''."""
    want = want_name.lower()
    for c in df.columns:
        if c.lower() == want:
            return c
    return ""

def apply_mapping_series(ser: pd.Series, mapping: dict, case_insensitive=True):
    """Áp mapping cho 1 Series (chuẩn hóa key trước)."""
    if case_insensitive:
        # chuẩn hóa key của mapping
        canon_map = {norm_text(k).lower(): v for k, v in mapping.items()}
        return ser.map(lambda v: canon_map.get(norm_text(v).lower(), v))
    else:
        canon_map = {norm_text(k): v for k, v in mapping.items()}
        return ser.map(lambda v: canon_map.get(norm_text(v), v))

def row_match(val, op, needle, ci):
    s = norm_text(val)
    n = norm_text(needle)
    if ci:
        s_cmp, n_cmp = s.lower(), n.lower()
    else:
        s_cmp, n_cmp = s, n

    if op == "equals":
        return s_cmp == n_cmp
    if op == "contains":
        return n_cmp in s_cmp
    if op == "startswith":
        return s_cmp.startswith(n_cmp)
    if op == "endswith":
        return s_cmp.endswith(n_cmp)
    if op == "regex":
        flags = re.IGNORECASE if ci else 0
        try:
            return re.search(n, s, flags=flags) is not None
        except re.error:
            return False
    if op == "inlist":
        items = [t.strip() for t in n.split("|") if t.strip()]
        if ci:
            items = [t.lower() for t in items]
            return s_cmp in items
        return s in items
    return False

def safe_sheet_name(name: str) -> str:
    return re.sub(r"[\\/*?:\\[\\]]", "_", name)[:31]

# ---- XỬ LÝ WORKBOOK ----
def main():
    xls = pd.ExcelFile(INPUT_XLSX)
    out_sheets = {}
    rule_idx_global = 1

    for sname in xls.sheet_names:
        df = pd.read_excel(INPUT_XLSX, sheet_name=sname)
        # 1) Replace theo mapping cho từng cột (nếu có mặt)
        col_manu = find_col(df, "Manufacturer")
        if col_manu:
            df[col_manu] = apply_mapping_series(df[col_manu], MAPPINGS_MANUFACTURER, case_insensitive=True)

        col_type = find_col(df, "Instrument Type")
        if col_type:
            df[col_type] = apply_mapping_series(df[col_type], MAPPINGS_INSTRUMENT_TYPE, case_insensitive=True)

        col_param = find_col(df, "Parameter/Method")
        if col_param:
            df[col_param] = apply_mapping_series(df[col_param], MAPPINGS_PARAMETER_METHOD, case_insensitive=True)

        # 2) Replace có điều kiện
        local_rule_counter = 1
        for r in RULES:
            if_col = find_col(df, r.get("if_col", ""))
            set_col = find_col(df, r.get("set_col", ""))
            if not if_col or not set_col:
                continue

            op   = r.get("op", "equals").lower()
            val  = r.get("value", "")
            ci   = bool(r.get("case_insensitive", True))
            audit = bool(r.get("audit", True))

            mask = df[if_col].map(lambda v: row_match(v, op, val, ci))

            if audit:
                old_col = f"{set_col}__old"
                if old_col not in df.columns:
                    df[old_col] = df[set_col]

            df.loc[mask, set_col] = r.get("set_value", "")

            if audit:
                flag_col = f"__rule_{rule_idx_global:02d}"
                text = f"{set_col}='{r.get('set_value','')}' if {r.get('if_col','')} {op} '{val}'"
                df[flag_col] = np.where(mask, text, df.get(flag_col, ""))

            local_rule_counter += 1
            rule_idx_global += 1

        out_sheets[sname] = df

    # 3) Lưu file
    with pd.ExcelWriter(OUTPUT_XLSX, engine="xlsxwriter") as writer:
        for sname, df in out_sheets.items():
            df.to_excel(writer, sheet_name=safe_sheet_name(sname), index=False)

    print(f"✅ Saved: {OUTPUT_XLSX}")

if __name__ == "__main__":
    main()
