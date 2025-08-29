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


def quote_sheet(name: str) -> str:
    # always quote sheet names for safety
    return f"'{name}'" if not (name.startswith("'") and name.endswith("'")) else name

def build_ext_vlookup(path: Path, sheet: str, table_range: str, key_ref: str, col_index: int) -> str:
    # 'C:\...\[file.xls]Sheet'!$B:$C  → VLOOKUP(key, that, col_index, FALSE)
    book = f"[{path.name}]"
    sheet_quoted = quote_sheet(sheet)
    xref = f"'{str(path.parent)}\\{book}{sheet_quoted[1:-1]}'!{table_range}"
    return f"VLOOKUP({key_ref},{xref},{col_index},FALSE)"

def copy_header_style(ws, src_col_idx: int, dst_col_idx: int):
    s = ws.cell(row=1, column=src_col_idx)
    d = ws.cell(row=1, column=dst_col_idx)
    d.value = s.value  # temp; caller will overwrite with new header text
    d.number_format = s.number_format
    from copy import copy as _cpy
    d.font = _cpy(s.font); d.fill = _cpy(s.fill); d.border = _cpy(s.border)
    d.alignment = _cpy(s.alignment); d.protection = _cpy(s.protection)
    # copy column width
    try:
        from openpyxl.utils import get_column_letter
        dst_letter = get_column_letter(dst_col_idx)
        src_letter = get_column_letter(src_col_idx)
        ws.column_dimensions[dst_letter].width = ws.column_dimensions[src_letter].width
    except Exception:
        pass

def copy_body_style_from_left(ws, row: int, col_idx: int):
    # style like the immediate left neighbor (common pattern in your sheets)
    if col_idx <= 1: 
        return
    s = ws.cell(row=row, column=col_idx - 1)
    d = ws.cell(row=row, column=col_idx)
    from copy import copy as _cpy
    d.number_format = s.number_format
    d.font = _cpy(s.font); d.fill = _cpy(s.fill); d.border = _cpy(s.border)
    d.alignment = _cpy(s.alignment); d.protection = _cpy(s.protection)


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

# --- related-party mapping (read .xls) ---
def normalize_id(x) -> str:
    if x is None:
        return ""
    s = str(x).strip()
    # keep leading zeros; if it looks like a float-ish "960286.0", strip trailing .0
    if s.endswith(".0") and s.replace(".", "", 1).isdigit():
        s = s[:-2]
    return s

def load_relparty_map(xls_path: Path) -> dict[str, str]:
    """
    Read first sheet of 關係企業(人).xls
    Col A = ID, Col C = Name. Returns {ID -> Name}.
    Requires pandas + xlrd (pip install pandas xlrd)
    """
    import pandas as pd
    df = pd.read_excel(xls_path, sheet_name=0, header=None, engine="xlrd")  # .xls
    # guard for short sheets
    if df.shape[1] < 3:
        return {}
    ids = df.iloc[:, 0].map(normalize_id)
    names = df.iloc[:, 2].fillna("").astype(str).str.strip()
    return {i: n for i, n in zip(ids, names) if i}

# Accounting format: 2-3 and 4-3
ACCOUNTING_FMT = "#,##0;[Red](#,##0);0;@"

def apply_accounting_format(ws, col_idx: int, start_row: int, end_row: int):
    for r in range(start_row, end_row + 1):
        ws.cell(row=r, column=col_idx).number_format = ACCOUNTING_FMT

# Appenders: 
def last_data_row(ws, key_col=1, max_gap=100):
    # scan down key_col (default A) until long empty gap
    rmax = ws.max_row or 1
    seen = 0
    for r in range(1, rmax + 1):
        v = ws.cell(row=r, column=key_col).value
        if v is not None and str(v).strip() != "":
            seen = r
    # if nothing found, still return 1 so headers can exist
    return max(seen, 1)


def append_calc_columns_23(ws, period: str, rates_path: Path, relparty_map: dict):
    from openpyxl.utils import column_index_from_string as colidx
    AK, AL, AM = colidx("AK"), colidx("AL"), colidx("AM")
    N, M, AJ = colidx("N"), colidx("M"), colidx("AJ")

    # headers
    copy_header_style(ws, AJ, AK); ws.cell(row=1, column=AK).value = "匯率"
    copy_header_style(ws, AJ, AL); ws.cell(row=1, column=AL).value = "換算台幣"
    copy_header_style(ws, AJ, AM); ws.cell(row=1, column=AM).value = "關係企業名稱"

    lr = last_data_row(ws, key_col=1)

    # formulas for rate / amount
    rates_vlk = lambda r: f"IF({ws.cell(row=r, column=N).coordinate}=\"NTD\",1," + \
        build_ext_vlookup(rates_path, "Summary", "$B:$C", ws.cell(row=r, column=N).coordinate, 2) + ")"

    for r in range(2, lr + 1):
        # 匯率 (AK)
        copy_body_style_from_left(ws, r, AK)
        ws.cell(row=r, column=AK).value = f"={rates_vlk(r)}"

        # 換算台幣 (AL) = M * AK
        copy_body_style_from_left(ws, r, AL)
        ws.cell(row=r, column=AL).value = f"={ws.cell(row=r, column=M).coordinate}*{ws.cell(row=r, column=AK).coordinate}"

        # 關係企業名稱 (AM) — resolve via mapping (no external formula)
        copy_body_style_from_left(ws, r, AM)
        key = normalize_id(ws.cell(row=r, column=AJ).value)
        ws.cell(row=r, column=AM).value = relparty_map.get(key, "")

    # accounting format for 換算台幣
    apply_accounting_format(ws, AL, 2, lr)


# def append_calc_columns_43(ws, period: str, rates_path: Path, relparty_path: Path):
#     """
#     Sheet: 4-3.應收關係人科餘
#     Existing data pasted to A:W → we append:
#       X: 匯率
#       Y: 換算台幣
#       Z: 關係企業名稱
#     Formulas:
#       匯率           = IF(K2="NTD",1, VLOOKUP(K2, [rates]Summary!$B:$C, 2, FALSE))
#       換算台幣       = L2*X2
#       關係企業名稱   = VLOOKUP(J2, [relparty]Sheet1!$A:$C, 3, FALSE)
#     """
#     from openpyxl.utils import column_index_from_string as colidx
#     X, Y, Z = colidx("X"), colidx("Y"), colidx("Z")
#     J, K, L = colidx("J"), colidx("K"), colidx("L")

#     # header styles copied from previous header (W)
#     from openpyxl.utils import column_index_from_string
#     W = column_index_from_string("W")
#     copy_header_style(ws, W, X); ws.cell(row=1, column=X).value = "匯率"
#     copy_header_style(ws, W, Y); ws.cell(row=1, column=Y).value = "換算台幣"
#     copy_header_style(ws, W, Z); ws.cell(row=1, column=Z).value = "關係企業名稱"

#     lr = last_data_row(ws, key_col=1)

#     rates_vlk = lambda r: f"IF({ws.cell(row=r, column=K).coordinate}=\"NTD\",1," + \
#         build_ext_vlookup(rates_path, "Summary", "$B:$C", ws.cell(row=r, column=K).coordinate, 2) + ")"
#     rel_vlk   = lambda r: build_ext_vlookup(relparty_path, "Sheet1", "$A:$C", ws.cell(row=r, column=J).coordinate, 3)

#     for r in range(2, lr + 1):
#         # 匯率
#         copy_body_style_from_left(ws, r, X)
#         ws.cell(row=r, column=X).value = f"={rates_vlk(r)}"
#         # 換算台幣
#         copy_body_style_from_left(ws, r, Y)
#         ws.cell(row=r, column=Y).value = f"={ws.cell(row=r, column=L).coordinate}*{ws.cell(row=r, column=X).coordinate}"
#         # 關係企業名稱
#         copy_body_style_from_left(ws, r, Z)
#         ws.cell(row=r, column=Z).value = f"={rel_vlk(r)}"


def append_calc_columns_43(ws, period: str, rates_path: Path, relparty_map: dict):
    from openpyxl.utils import column_index_from_string as colidx
    X, Y, Z = colidx("X"), colidx("Y"), colidx("Z")
    J, K, L = colidx("J"), colidx("K"), colidx("L")
    W = colidx("W")

    # headers
    copy_header_style(ws, W, X); ws.cell(row=1, column=X).value = "匯率"
    copy_header_style(ws, W, Y); ws.cell(row=1, column=Y).value = "換算台幣"
    copy_header_style(ws, W, Z); ws.cell(row=1, column=Z).value = "關係企業名稱"

    lr = last_data_row(ws, key_col=1)

    rates_vlk = lambda r: f"IF({ws.cell(row=r, column=K).coordinate}=\"NTD\",1," + \
        build_ext_vlookup(rates_path, "Summary", "$B:$C", ws.cell(row=r, column=K).coordinate, 2) + ")"

    for r in range(2, lr + 1):
        # 匯率 (X)
        copy_body_style_from_left(ws, r, X)
        ws.cell(row=r, column=X).value = f"={rates_vlk(r)}"

        # 換算台幣 (Y) = L * X
        copy_body_style_from_left(ws, r, Y)
        ws.cell(row=r, column=Y).value = f"={ws.cell(row=r, column=L).coordinate}*{ws.cell(row=r, column=X).coordinate}"

        # 關係企業名稱 (Z) — resolve via mapping
        copy_body_style_from_left(ws, r, Z)
        key = normalize_id(ws.cell(row=r, column=J).value)
        ws.cell(row=r, column=Z).value = relparty_map.get(key, "")

    # accounting format for 換算台幣
    apply_accounting_format(ws, Y, 2, lr)



# ---------- CLI ----------
# def main():
#     parser = argparse.ArgumentParser(description="Fill updated YTM forms by direct copy/paste with styles.")
#     parser.add_argument("--template", required=True, help="Path to the template workbook (will be copied unless --inplace).")
#     parser.add_argument("--out", help="Output path (.xlsx). If omitted, writes to default unless --inplace.")
#     parser.add_argument("--inplace", action="store_true", help="Overwrite template in-place.")

#     parser.add_argument("--task", required=True,
#                         choices=["copy_4_3", "copy_2_3", "both"],
#                         help="Which sheet(s) to fill.")
#     parser.add_argument("--src-43", default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\export_關係人交易-應收帳款.xlsx",
#                         help="Source workbook for 4-3 (columns B:X).")
#     parser.add_argument("--src-23", default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\export_關係人交易-收入.xlsx",
#                         help="Source workbook for 2-3 (columns A:AJ).")

#     args = parser.parse_args()

#     template_path = Path(args.template)
#     if not template_path.exists():
#         raise FileNotFoundError(f"Template not found: {template_path}")

#     out_path = Path(args.out) if args.out else (
#     template_path if args.inplace else OUTPUT_DIR / f"copy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
#     )


#     final_out = load_output_from_template(template_path, out_path, args.inplace)

#     if args.task in ("copy_4_3", "both"):
#         copy_43(template_path, Path(args.src_43), final_out)

#     if args.task in ("copy_2_3", "both"):
#         copy_23(template_path, Path(args.src_23), final_out)

#     print(f"Done → {final_out}")

# ---------- CLI ----------
def main():
    parser = argparse.ArgumentParser(description="Fill updated YTM forms by direct copy/paste with styles.")
    parser.add_argument("--template", required=True, help="Path to the template workbook (will be copied unless --inplace).")
    parser.add_argument("--out", help="Output path (.xlsx). If omitted, writes to default unless --inplace.")
    parser.add_argument("--inplace", action="store_true", help="Overwrite template in-place.")

    parser.add_argument("--task", required=True,
                        choices=["copy_4_3", "copy_2_3", "both"],
                        help="Which sheet(s) to fill.")
    parser.add_argument("--src-43", help="Source workbook for 4-3 (columns B:X).")
    parser.add_argument("--src-23", help="Source workbook for 2-3 (columns A:AJ).")
    # parser.add_argument("--src-43", default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\export_關係人交易-應收帳款.xlsx",
    #                     help="Source workbook for 4-3 (columns B:X).")
    # parser.add_argument("--src-23", default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\export_關係人交易-收入.xlsx",
    #                     help="Source workbook for 2-3 (columns A:AJ).")

    # NEW: external lookups
    parser.add_argument("--period", required=True, help="e.g., 202504")
    # parser.add_argument("--rates-path",
    #                     help="External rates workbook path; overrides default pattern: <period> Ending 及 Avg (資通版本).xls")
    # parser.add_argument("--relparty-path",
    #                     default=r"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\關係企業(人).xls",
    #                     help="External related-party master workbook path (default: 關係企業(人).xls)")
    parser.add_argument("--rates-path", help="External rates workbook path; overrides default pattern.")
    parser.add_argument("--relparty-path", help="External related-party master workbook path.")

    args = parser.parse_args()
    # Base folder inside the repo
    BASE_TPL = PROJECT_ROOT / "ytm_forms" / "data" / "template" / "關係人"

    # Resolve sources (allow override via CLI)
    src_43 = Path(args.src_43) if args.src_43 else BASE_TPL / "export_關係人交易-應收帳款.xlsx"
    src_23 = Path(args.src_23) if args.src_23 else BASE_TPL / "export_關係人交易-收入.xlsx"

    # Resolve external files (allow override via CLI)
    yyyymm = args.period
    default_rates = BASE_TPL / f"{yyyymm} Ending 及 Avg (資通版本).xls"
    rates_path   = Path(args.rates_path) if args.rates_path else default_rates
    relparty_path = Path(args.relparty_path) if args.relparty_path else (BASE_TPL / "關係企業(人).xls")


    # validate template
    template_path = Path(args.template)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    # resolve output path to ytm_forms/data/output (anchored to project root)
    out_path = Path(args.out) if args.out else (
        template_path if args.inplace else OUTPUT_DIR / f"copy_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    final_out = load_output_from_template(template_path, out_path, args.inplace)

    if args.task in ("copy_4_3", "both"):
        copy_43(template_path, src_43, final_out)
    if args.task in ("copy_2_3", "both"):
        copy_23(template_path, src_23, final_out)

    # warnings if links missing (optional)
    if not rates_path.exists():
        print(f"[WARN] Rates workbook not found: {rates_path}")
    if not relparty_path.exists():
        print(f"[WARN] Related-party workbook not found: {relparty_path}")

    # build {ID -> Name} map once
    relparty_map = load_relparty_map(relparty_path)


    # append 3 columns (匯率、換算台幣、關係企業名稱) with external formulas
    wb_final = load_workbook(final_out, data_only=False)

    if "4-3.應收關係人科餘" in wb_final.sheetnames:
        append_calc_columns_43(wb_final["4-3.應收關係人科餘"], args.period, rates_path, relparty_map)

    if "2-3.銷貨明細" in wb_final.sheetnames:
        append_calc_columns_23(wb_final["2-3.銷貨明細"], args.period, rates_path, relparty_map)


    wb_final.save(final_out)
    wb_final.close()

    print(f"Done → {final_out}")


if __name__ == "__main__":
    main()
