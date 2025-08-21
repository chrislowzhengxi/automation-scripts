#!/usr/bin/env python3
"""
Group export rows by G/L account using a mapping file and create sub-sheets
named "number name" with columns B..X from the export.

Default behavior:
- Reads the first sheet of the export workbook (unless --sheet is specified).
- Finds the "G/L科目" column in the export (required).
- Reads mapping from the first sheet of the mapping file: col A=number, col B=name.
- Writes a NEW workbook next to the export with suffix "_grouped.xlsx".
  (Use --inplace to modify the export file directly.)

Usage examples (Windows):
  py group_by_gl.py ^
    --export ".\\export-科餘-1000-asset.xlsx" ^
    --mapping ".\\會計科目對照表.xlsx"

  # Pick a specific export sheet by name:
  py group_by_gl.py --export ".\\export-科餘-1000-asset.xlsx" --mapping ".\\會計科目對照表.xlsx" --sheet "Sheet1"

  # Modify original in place (adds sub-sheets into the export file):
  py group_by_gl.py --export ".\\export-科餘-1000-asset.xlsx" --mapping ".\\會計科目對照表.xlsx" --inplace
"""

from __future__ import annotations
import argparse
from pathlib import Path
import re
import sys
import pandas as pd
import openpyxl


def norm_code(x):
    """Normalize account code to a clean string (strip, remove commas, drop trailing .0)."""
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = s.replace(",", "").strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s or None


def sanitize_sheet_name(name: str) -> str:
    """Excel sheet name rules: <=31 chars, cannot contain : \ / ? * [ ]"""
    name = re.sub(r'[:\\/\?\*\[\]]', ' ', name)
    return name[:31] if len(name) > 31 else name


def choose_output_path(export_path: Path, inplace: bool, output: str | None) -> Path:
    if inplace:
        return export_path
    if output:
        return Path(output)
    return export_path.with_name(export_path.stem + "_grouped.xlsx")


def read_export_frame(export_path: Path, sheet_name: str | None) -> tuple[pd.DataFrame, str]:
    # Detect first sheet if not given
    if sheet_name is None:
        xls = pd.ExcelFile(export_path)
        sheet_name = xls.sheet_names[0]
    df = pd.read_excel(export_path, sheet_name=sheet_name, dtype=object)
    # Normalize headers
    df.columns = [str(c).strip() for c in df.columns]
    return df, sheet_name


def find_gl_column(df: pd.DataFrame) -> str:
    candidates = [c for c in df.columns if c.strip() == "G/L科目"]
    if not candidates:
        raise ValueError("Could not find column 'G/L科目' in the export.")
    return candidates[0]


def load_mapping(mapping_path: Path) -> dict[str, str]:
    # Read first sheet, columns A & B (0,1) only
    df_map_raw = pd.read_excel(mapping_path, sheet_name=0, dtype=object, usecols=[0, 1], header=0)
    # Standardize column names to "number" and "name"
    df_map_raw = df_map_raw.rename(columns={df_map_raw.columns[0]: "number", df_map_raw.columns[1]: "name"})
    df_map_raw["number_norm"] = df_map_raw["number"].apply(norm_code)
    df_map = df_map_raw.dropna(subset=["number_norm"]).copy()
    # Build mapping dict: normalized number -> display name (string)
    return dict(zip(df_map["number_norm"], df_map["name"].fillna("").astype(str)))


def pick_columns_B_to_X(df: pd.DataFrame) -> list[str]:
    # Columns are 0-based indices; B..X = 1..23 (inclusive) -> slice [1:24]
    bx_end = min(24, len(df.columns))
    cols = [c for c in df.columns[1:bx_end] if c != "_code"]
    return cols


def ensure_unique_title(base_title: str, taken: set[str]) -> str:
    title = sanitize_sheet_name(base_title) or "Sheet"
    if title not in taken:
        return title
    # If conflict, add suffixes
    suffix = 1
    while True:
        candidate = sanitize_sheet_name(f"{title[:25]}_{suffix}")
        if candidate not in taken:
            return candidate
        suffix += 1


def group_export_by_account(
    export_path: Path,
    mapping_path: Path,
    output_path: Path,
    sheet_name: str | None,
    inplace: bool,
) -> dict:
    # 1) Read export and identify G/L column
    df_export, sheet_used = read_export_frame(export_path, sheet_name)
    gl_col = find_gl_column(df_export)

    # 2) Load mapping
    number_to_name = load_mapping(mapping_path)

    # 3) Normalize account code and filter valid rows (skip empties)
    df_export["_code"] = df_export[gl_col].apply(norm_code)
    df_export_valid = df_export[df_export["_code"].notna()].copy()
    df_export_valid = df_export_valid[df_export_valid["_code"].isin(number_to_name.keys())]

    # 4) Choose columns B..X
    selected_cols = pick_columns_B_to_X(df_export)

    # 5) Load workbook (we either edit the export file directly or write a copy)
    if inplace:
        wb = openpyxl.load_workbook(export_path)
    else:
        # Start from a fresh copy: copy the original sheets & data
        wb = openpyxl.load_workbook(export_path)

    # Track existing sheet names to avoid collisions
    used_titles = {ws.title for ws in wb.worksheets}

    # 6) For each unique account code, create sub-sheet and paste header + rows (B..X)
    for code, grp in df_export_valid.groupby("_code"):
        name = number_to_name.get(code, "").strip()
        base_title = f"{code} {name}".strip()
        title = ensure_unique_title(base_title, used_titles)
        used_titles.add(title)

        ws = wb.create_sheet(title=title)

        # Header
        for j, col_name in enumerate(selected_cols, start=1):
            ws.cell(row=1, column=j, value=str(col_name))

        # Rows
        sub = grp[selected_cols]
        for i, row in enumerate(sub.itertuples(index=False, name=None), start=2):
            for j, val in enumerate(row, start=1):
                ws.cell(row=i, column=j, value=val)

    # 7) Save
    if inplace:
        wb.save(export_path)
        saved_to = str(export_path)
    else:
        wb.save(output_path)
        saved_to = str(output_path)

    return {
        "export_sheet": sheet_used,
        "gl_col": gl_col,
        "unique_accounts": int(df_export_valid["_code"].nunique()),
        "rows_grouped": int(len(df_export_valid)),
        "saved_to": saved_to,
        "columns_B_to_X": selected_cols,
    }


def main():
    p = argparse.ArgumentParser(description="Group export rows by G/L科目 and create sub-sheets per account.")
    p.add_argument("--export", required=False, default="export-科餘-1000-asset.xlsx",
                   help="Path to the export workbook (default: ./export-科餘-1000-asset.xlsx)")
    p.add_argument("--mapping", required=False, default="會計科目對照表.xlsx",
                   help="Path to the mapping workbook (default: ./會計科目對照表.xlsx)")
    p.add_argument("--output", required=False, default=None,
                   help="Output path for the new workbook (ignored if --inplace). Defaults to <export>_grouped.xlsx")
    p.add_argument("--sheet", required=False, default=None,
                   help="Export sheet name to read. If omitted, uses the first sheet.")
    p.add_argument("--inplace", action="store_true",
                   help="Modify the export file in place (adds sub-sheets into the same workbook).")
    args = p.parse_args()

    export_path = Path(args.export).expanduser().resolve()
    mapping_path = Path(args.mapping).expanduser().resolve()

    if not export_path.exists():
        print(f"[ERROR] Export file not found: {export_path}", file=sys.stderr)
        sys.exit(2)
    if not mapping_path.exists():
        print(f"[ERROR] Mapping file not found: {mapping_path}", file=sys.stderr)
        sys.exit(2)

    output_path = choose_output_path(export_path, args.inplace, args.output)

    stats = group_export_by_account(
        export_path=export_path,
        mapping_path=mapping_path,
        output_path=output_path,
        sheet_name=args.sheet,
        inplace=args.inplace,
    )

    print("[OK] Grouping complete.")
    for k, v in stats.items():
        print(f"- {k}: {v}")


if __name__ == "__main__":
    main()
