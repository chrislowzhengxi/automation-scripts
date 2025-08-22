#!/usr/bin/env python3
import argparse, os, shutil
from datetime import datetime
from pathlib import Path
from copy import copy
from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter


# Base = ytm_forms/
BASE_DIR = Path(__file__).resolve().parents[1]
DEFAULT_OUT_DIR = BASE_DIR / "data" / "output"

PROJECT_ROOT = Path(__file__).resolve().parents[2]
OUTPUT_DIR = PROJECT_ROOT / "ytm_forms" / "data" / "output"

# ---------- helpers ----------
def ensure_parent(p: Path):
    p.parent.mkdir(parents=True, exist_ok=True)

def copy_cell_full(src_cell, dst_cell):
    dst_cell.value = src_cell.value
    dst_cell.number_format = src_cell.number_format
    dst_cell.font = copy(src_cell.font)
    dst_cell.fill = copy(src_cell.fill)
    dst_cell.border = copy(src_cell.border)
    dst_cell.alignment = copy(src_cell.alignment)
    dst_cell.protection = copy(src_cell.protection)

def copy_col_widths(ws_src, ws_dst, c1, c2, dst_start_col):
    for j in range(c2 - c1 + 1):
        src_col_idx = c1 + j
        dst_col_idx = dst_start_col + j
        src_letter = get_column_letter(src_col_idx)
        dst_letter = get_column_letter(dst_col_idx)
        try:
            ws_dst.column_dimensions[dst_letter].width = ws_src.column_dimensions[src_letter].width
        except Exception:
            pass

def copy_row_heights(ws_src, ws_dst, r1, r2, dst_start_row):
    for i in range(r2 - r1 + 1):
        src_row = r1 + i
        dst_row = dst_start_row + i
        try:
            ws_dst.row_dimensions[dst_row].height = ws_src.row_dimensions[src_row].height
        except Exception:
            pass

def clear_sheet(ws):
    ws.delete_rows(1, ws.max_row if ws.max_row else 1)

def copy_block(ws_src, ws_dst, r1, r2, c1, c2, dst_row, dst_col):
    # copy values/styles
    for i in range(r2 - r1 + 1):
        for j in range(c2 - c1 + 1):
            s = ws_src.cell(row=r1 + i, column=c1 + j)
            d = ws_dst.cell(row=dst_row + i, column=dst_col + j)
            copy_cell_full(s, d)
    # aesthetics
    copy_row_heights(ws_src, ws_dst, r1, r2, dst_row)
    copy_col_widths(ws_src, ws_dst, c1, c2, dst_col)

def load_output_from_template(template_path: Path, out_path: Path, inplace: bool) -> Path:
    if inplace:
        return template_path
    ensure_parent(out_path)
    shutil.copy(template_path, out_path)
    return out_path

# ---------- tasks ----------
def copy_43(template_path: Path, src_43: Path, out_path: Path, sheet_name="4-3.應收關係人科餘"):
    # Source: columns B..X → Destination: paste at B1 (keep alignment with template)
    c1, c2 = column_index_from_string("B"), column_index_from_string("X")
    dst_col = column_index_from_string("A")
    dst_row = 1

    wb_src = load_workbook(src_43, data_only=False)
    ws_src = wb_src.active
    r1, r2 = 1, ws_src.max_row or 1

    wb_dst = load_workbook(out_path)
    ws_dst = wb_dst[sheet_name] if sheet_name in wb_dst.sheetnames else wb_dst.active

    clear_sheet(ws_dst)  # as you said, we start with empty 4-3 each run
    copy_block(ws_src, ws_dst, r1, r2, c1, c2, dst_row, dst_col)

    wb_dst.save(out_path)
    wb_dst.close()
    wb_src.close()

def copy_23(template_path: Path, src_23: Path, out_path: Path, sheet_name="2-3.銷貨明細"):
    # Source: columns A..AJ → Destination: paste at A1
    c1, c2 = column_index_from_string("A"), column_index_from_string("AJ")
    dst_col = column_index_from_string("A")
    dst_row = 1

    wb_src = load_workbook(src_23, data_only=False)
    ws_src = wb_src.active
    r1, r2 = 1, ws_src.max_row or 1

    wb_dst = load_workbook(out_path)
    ws_dst = wb_dst[sheet_name] if sheet_name in wb_dst.sheetnames else wb_dst.active

    clear_sheet(ws_dst)  # start clean for 2-3 as well
    copy_block(ws_src, ws_dst, r1, r2, c1, c2, dst_row, dst_col)

    wb_dst.save(out_path)
    wb_dst.close()
    wb_src.close()

# ---------- CLI ----------
def main():
    parser = argparse.ArgumentParser(description="Fill updated YTM forms by direct copy/paste with styles.")
    parser.add_argument("--template", required=True, help="Path to the template workbook (will be copied unless --inplace).")
    parser.add_argument("--out", help="Output path (.xlsx). If omitted, writes to default unless --inplace.")
    parser.add_argument("--inplace", action="store_true", help="Overwrite template in-place.")

    parser.add_argument("--task", required=True,
                        choices=["copy_4_3", "copy_2_3", "both"],
                        help="Which sheet(s) to fill.")
    parser.add_argument("--src-43", default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\export_關係人交易-應收帳款.xlsx",
                        help="Source workbook for 4-3 (columns B:X).")
    parser.add_argument("--src-23", default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\export_關係人交易-收入.xlsx",
                        help="Source workbook for 2-3 (columns A:AJ).")

    args = parser.parse_args()

    template_path = Path(args.template)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    out_path = Path(args.out) if args.out else (
    template_path if args.inplace else OUTPUT_DIR / f"copy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )


    final_out = load_output_from_template(template_path, out_path, args.inplace)

    if args.task in ("copy_4_3", "both"):
        copy_43(template_path, Path(args.src_43), final_out)

    if args.task in ("copy_2_3", "both"):
        copy_23(template_path, Path(args.src_23), final_out)

    print(f"Done → {final_out}")

if __name__ == "__main__":
    main()
