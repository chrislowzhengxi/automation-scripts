#!/usr/bin/env python3
import argparse, os, re, shutil
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell
from copy import copy

src = Path(r"C:\Users\TP2507088\Downloads\202504YTM WITS-C 上櫃公司與關係人間重要交易資訊.xlsx")
dst = src.with_name(f"Copy {src.name}")

shutil.copy(src, dst)
print(f"Copied to: {dst}")

TARGET_SHEET = "2-2.銷貨倍力"
COLUMN_A = 1
COLUMN_S = 19
PERIOD_RX = re.compile(r"(20\d{2})(0[1-9]|1[0-2])")

# ---------- common helpers ----------
def find_downloads() -> Path:
    up = os.environ.get("USERPROFILE")
    d = Path(up) / "Downloads" if up else Path.home() / "Downloads"
    return d

def pick_file_by_period(prefix: str, period: str, explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise FileNotFoundError(f"Path not found: {p}")
        return p

    downloads = find_downloads()
    candidates = sorted(Path(downloads).glob(f"*{prefix}*.xlsx"))
    if not candidates:
        raise FileNotFoundError(f"No {prefix} files in {downloads}")

    # 1) exact period match
    period_matches = [p for p in candidates if period in p.name]
    if period_matches:
        return max(period_matches, key=lambda p: p.stat().st_mtime)

    # 2) else pick highest YYYYMM we can parse
    def pnum(p: Path):
        m = PERIOD_RX.search(p.name)
        return int(m.group(1) + m.group(2)) if m else None

    with_period = [(p, pnum(p)) for p in candidates]
    with_period = [(p, n) for (p, n) in with_period if n is not None]
    if with_period:
        max_period = max(n for _, n in with_period)
        newest_in_max = max([p for (p, n) in with_period if n == max_period],
                            key=lambda p: p.stat().st_mtime)
        return newest_in_max

    # 3) fallback newest by mtime
    return max(candidates, key=lambda p: p.stat().st_mtime)

def copy_cell(src, dst):
    # value
    dst.value = src.value
    # number format + basic style so red/parentheses/accounting look right
    dst.number_format = src.number_format
    dst.font = copy(src.font)
    dst.fill = copy(src.fill)
    dst.border = copy(src.border)
    dst.alignment = copy(src.alignment)
    dst.protection = copy(src.protection)

def dest_anchor(ws, r, c):
    for rng in ws.merged_cells.ranges:
        if rng.min_row <= r <= rng.max_row and rng.min_col <= c <= rng.max_col:
            return rng.min_row, rng.min_col
    return r, c

def default_export_path(explicit: str | None) -> Path:
    if explicit:
        p = Path(explicit)
        if not p.exists():
            raise FileNotFoundError(f"export.xlsx not found: {p}")
        return p
    d = find_downloads()
    p = d / "export.xlsx"
    if not p.exists():
        raise FileNotFoundError(f"export.xlsx not found at {p}")
    return p


# ---------- task: MRS0014 → 2-2 ----------
def run_mrs0014(template_path: Path, period: str, mrs_path: str | None, out_path: Path):
    mrs_file = pick_file_by_period("MRS0014", period, mrs_path)

    # read values for 421007 / 421807 from column S
    wb_src = load_workbook(mrs_file, data_only=True)
    vals = {"421007": 0, "421807": 0}
    try:
        for ws in wb_src.worksheets:
            for r in range(1, (ws.max_row or 5000) + 1):
                a_val = ws.cell(row=r, column=COLUMN_A).value
                if a_val is None:
                    continue
                key = str(a_val).strip()
                if key in vals:
                    s_val = ws.cell(row=r, column=COLUMN_S).value or 0
                    try:
                        vals[key] += s_val
                    except TypeError:
                        pass
    finally:
        wb_src.close()

    # write to template (C18, C19, sum to C21)
    wb = load_workbook(template_path)
    try:
        ws = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb.active
        ws["C18"] = vals["421007"]
        ws["C19"] = vals["421807"]
        ws["C21"] = "=SUM(C18:C19)"
        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
    finally:
        wb.close()

# ---------- task: RPTIS10 → 2-2 ----------
def run_rptis10(template_path: Path, period: str, rpt_path: str | None, out_path: Path,
                src_sheet: str | None = None, src_rows=(6,11), src_cols=(1,3),
                dst_start_row=7, dst_start_col=1):
    rpt_file = pick_file_by_period("RPTIS10", period, rpt_path)

    wb_src = load_workbook(rpt_file, data_only=False)  # keep styles
    try:
        ws_src = wb_src[src_sheet] if (src_sheet and src_sheet in wb_src.sheetnames) else wb_src.active
        r1, r2 = src_rows
        c1, c2 = src_cols
        block = []
        for r in range(r1, r2+1):
            row_cells = []
            for c in range(c1, c2+1):
                row_cells.append(ws_src.cell(row=r, column=c))
            block.append(row_cells)
    finally:
        wb_src.close()

    wb = load_workbook(template_path)
    try:
        ws_dst = wb[TARGET_SHEET] if TARGET_SHEET in wb.sheetnames else wb.active
        # for i, row_cells in enumerate(block):
        #     dst_row = dst_start_row + i
        #     for j, src_cell in enumerate(row_cells):
        #         dst_col = dst_start_col + j
        #         copy_cell(src_cell, ws_dst.cell(row=dst_row, column=dst_col))
        anchors_written = set()
        for i, row_cells in enumerate(block):
            dst_row = dst_start_row + i
            for j, src_cell in enumerate(row_cells):
                dst_col = dst_start_col + j
                ar, ac = dest_anchor(ws_dst, dst_row, dst_col)
                key = (ar, ac)
                if key in anchors_written:
                    continue
                anchors_written.add(key)
                dst_cell = ws_dst.cell(row=ar, column=ac)
                copy_cell(src_cell, dst_cell)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(out_path)
    finally:
        wb.close()


# ------- 4.3 -----------
# def find_reserved_header_row(ws, header_keywords=("公司代碼", "年度/月分")):
#     for r in range(1, ws.max_row + 1):
#         row_vals = [str(ws.cell(row=r, column=c).value or "") for c in range(1, ws.max_column + 1)]
#         if any(keyword in cell for cell in row_vals for keyword in header_keywords):
#             return r
#     return None
def find_header_rows(ws, header_keywords=("公司代碼", "年度/月分"), min_row=5):
    """Return a sorted list of row numbers that look like the 4-3 header."""
    rows = []
    for r in range(min_row, ws.max_row + 1):
        row_vals = [str(ws.cell(row=r, column=c).value or "") for c in range(1, ws.max_column + 1)]
        if any((kw in cell) for cell in row_vals for kw in header_keywords):
            rows.append(r)
    return rows



def run_export_paste(template_path: Path,
                     export_path: Path,
                     out_path: Path,
                     dest_sheet: str,
                     src_sheet: str | None = None,
                     src_cols=("B","W"),
                     dst_cols=("A","V"),
                     start_row=2,
                     dst_start_row=10,
                     buffer_rows=23):
    src_wb = load_workbook(export_path, data_only=True)
    try:
        ws_src = src_wb[src_sheet] if (src_sheet and src_sheet in src_wb.sheetnames) else src_wb.active
        c1 = ws_src[src_cols[0]+"1"].column
        c2 = ws_src[src_cols[1]+"1"].column

        # find last non-empty row in export.xlsx
        last_src_row = start_row - 1
        for r in range(start_row, ws_src.max_row + 1):
            if any(ws_src.cell(row=r, column=c).value is not None for c in range(c1, c2+1)):
                last_src_row = r
        if last_src_row < start_row:
            raise ValueError("No data rows found in export.xlsx (B..W)")

        dst_wb = load_workbook(template_path)
        try:
            ws_dst = dst_wb[dest_sheet] if dest_sheet in dst_wb.sheetnames else dst_wb.active

            # detect header row dynamically
            # reserved_header_row = find_reserved_header_row(ws_dst)
            # if reserved_header_row is None:
            #     raise ValueError("Could not find reserved header row in destination sheet.")

            # allowed_last_row = reserved_header_row - buffer_rows - 1
            # max_rows_allowed = allowed_last_row - dst_start_row + 1
            # rows_to_write = last_src_row - start_row + 1

            # if rows_to_write > max_rows_allowed:
            #     raise ValueError(
            #         f"Too many rows ({rows_to_write}) to paste. "
            #         f"Only {max_rows_allowed} fit between row {dst_start_row} "
            #         f"and reserved header row {reserved_header_row} with buffer {buffer_rows}."
            # #     )
            # reserved_header_row = find_reserved_header_row(ws_dst, min_row=50)

            # if reserved_header_row is not None:
            #     allowed_last_row = reserved_header_row - buffer_rows - 1
            #     max_rows_allowed = allowed_last_row - dst_start_row + 1
            #     rows_to_write = last_src_row - start_row + 1

            #     if max_rows_allowed < 0:
            #         raise ValueError(
            #             f"Reserved header detected too close to start (header at row {reserved_header_row}, "
            #             f"buffer {buffer_rows}). Increase dst_start_row or reduce buffer."
            #         )

            #     if rows_to_write > max_rows_allowed:
            #         raise ValueError(
            #             f"Too many rows ({rows_to_write}) to paste. Only {max_rows_allowed} fit between "
            #             f"row {dst_start_row} and reserved header row {reserved_header_row} with buffer {buffer_rows}."
            #         )
            # detect top and reserved header rows
            header_rows = find_header_rows(ws_dst, min_row=5)

            if not header_rows:
                # no header detected: keep your current dst_start_row and skip reserve check
                top_header = None
                reserved_header_row = None
            else:
                top_header = header_rows[0]            # e.g., 9
                reserved_header_row = header_rows[1] if len(header_rows) >= 2 else (top_header + 23)  # e.g., 32
                # force paste to start below the top header
                dst_start_row = max(dst_start_row, top_header + 1)

            # capacity check only if we have a reserved header
            if reserved_header_row is not None:
                allowed_last_row = reserved_header_row - buffer_rows - 1  # keep 23 blank rows
                max_rows_allowed = allowed_last_row - dst_start_row + 1
                rows_to_write = last_src_row - start_row + 1

                if max_rows_allowed < 0:
                    raise ValueError(
                        f"Reserved header at row {reserved_header_row} too close to start "
                        f"(dst_start_row={dst_start_row}, buffer={buffer_rows})."
                    )
                if rows_to_write > max_rows_allowed:
                    raise ValueError(
                        f"Too many rows ({rows_to_write}) to paste. Only {max_rows_allowed} fit between "
                        f"row {dst_start_row} and reserved header row {reserved_header_row} with buffer {buffer_rows}."
                    )


            dst_c1 = ws_dst[dst_cols[0]+"1"].column
            for r in range(start_row, last_src_row + 1):
                dst_row = dst_start_row + (r - start_row)
                for src_c in range(c1, c2+1):
                    val = ws_src.cell(row=r, column=src_c).value
                    dst_c = dst_c1 + (src_c - c1)
                    cell = ws_dst.cell(row=dst_row, column=dst_c)
                    cell.value = val

                    if src_c in (13, 15) and isinstance(val, (int, float)):
                        cell.number_format = '#,##0.00' if isinstance(val, float) and abs(val - int(val)) > 1e-9 else '#,##0'

            out_path.parent.mkdir(parents=True, exist_ok=True)
            dst_wb.save(out_path)
        finally:
            dst_wb.close()
    finally:
        src_wb.close()



# ---------- CLI ----------
def main():
    ap = argparse.ArgumentParser(description="Fill YTM forms")
    ap.add_argument("--task", required=True,
                    choices=["mrs0014", "rptis10", "both", "export_4_3"])
    ap.add_argument("--period", required=True, help="e.g., 202504")
    ap.add_argument("--template", required=True, help="Path to the template workbook to fill")
    ap.add_argument("--out", help="Output path (.xlsx). Omit and use --inplace to overwrite template")
    ap.add_argument("--dest-start-row", type=int, default=10,
                help="First row to paste into on destination sheet (export_4_3)")
    ap.add_argument("--inplace", action="store_true", help="Overwrite the template file")

    # 2-2 options
    ap.add_argument("--mrs", help="Explicit path to MRS0014 (optional)")
    ap.add_argument("--rptis", help="Explicit path to RPTIS10 (optional)")
    ap.add_argument("--rpt-source-sheet", help="RPTIS10 source sheet name (optional)")

    # 4-3 options
    ap.add_argument("--export", help="Path to export.xlsx (optional; defaults to Downloads/export.xlsx)")
    ap.add_argument("--dest-sheet", default="4-3", help="Destination sheet name for the export paste")

    args = ap.parse_args()

    template_path = Path(args.template)
    if not template_path.exists():
        raise FileNotFoundError(f"Template not found: {template_path}")

    out_path = Path(args.out) if args.out else (
        template_path if args.inplace else Path(f"ytm_forms/data/output/{args.task}_{args.period}.xlsx")
    )

    if args.task == "mrs0014":
        run_mrs0014(template_path, args.period, args.mrs, out_path)
    elif args.task == "rptis10":
        run_rptis10(template_path, args.period, args.rptis, out_path, src_sheet=args.rpt_source_sheet)
    elif args.task == "both":
        run_mrs0014(template_path, args.period, args.mrs, out_path)
        run_rptis10(template_path, args.period, args.rptis, out_path, src_sheet=args.rpt_source_sheet)
    else:  # export_4_3
        exp_path = default_export_path(args.export)
        run_export_paste(template_path, exp_path, out_path,
                        dest_sheet=args.dest_sheet, dst_start_row=args.dest_start_row)

    print(f"Done → {out_path}")

if __name__ == "__main__":
    main()

