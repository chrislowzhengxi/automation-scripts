from collections import defaultdict
import shutil
from pathlib import Path
from parsers import CitiParser, CTBCParser, MegaParser, FubonParser, SinopacParser, ESunParser, BankParserBase
from fuzzy_matcher import match_entries_interactive, match_entries_debug
from utils import log_skipped

PARSER_REGISTRY = {
    "花旗": CitiParser,
    "中信": CTBCParser,
    "兆豐": MegaParser,
    "富邦": FubonParser,
    "永豐": SinopacParser,
    "玉山": ESunParser,
    # …more banks later…
}

def make_parser(path: Path) -> BankParserBase:
    stem = path.stem
    for key, cls in PARSER_REGISTRY.items():
        if key in stem:
            return cls(path)
    raise RuntimeError(f"No parser registered for {path.name}")


import argparse
import pandas as pd
import openpyxl
from datetime import datetime
from openpyxl.styles import Font
from rapidfuzz import process, fuzz
from pathlib import Path

# ─────────────── Configuration ───────────────
BASE_DIR        = Path("~/Downloads/Banks").expanduser()
BANK_FILE       = BASE_DIR / "花旗銀行對帳單-20250625.xlsx"
BANK_SHEET      = "Sheet2"
DB_FILE         = BASE_DIR / "會計憑證導入模板 - 1000 客戶資料庫.xls"
DB_SHEET        = "客戶資料庫"
FUZZY_THRESHOLD = 80
# OUTPUT_FILE = BASE_DIR / "會計憑證導入模板 - 空白檔案.xlsx"
RED_FONT    = Font(color="FF0000")


TEMPLATE_FILE = BASE_DIR / "會計憑證導入模板 - 空白檔案.xlsx"
def daily_output_path(post_date: str) -> Path:
    # post_date like "20250715"
    return BASE_DIR / f"會計憑證導入模板 - {post_date}.xlsx"

COL_DESC = "E"
COL_AMT  = "G"
KEYWORD  = "細節描述"

BANK_MAP = {
    "花旗": "花旗營業 NTD 0005",
    "中信": "中信營業 NTD 0800",
    "兆豐": "兆豐竹科新安 NTD 2656",
    "富邦": "富邦仁愛 NTD 6332",
    "永豐": "永豐城中 NTD 7978",
    "玉山": "玉山營業 NTD 8563",
    # …add more banks here…
}

# ─────────────── Functions ───────────────

def parse_args():
    p = argparse.ArgumentParser()
    p.add_argument(
        "--file", "-f",
        required=True,
        help="Path to the bank statement Excel file (xls or xlsx)"
    )
    p.add_argument("--date", "-d",
                   help="Posting date in YYYYMMDD (defaults to today)")
    return p.parse_args()

def detect_bank(stem, bank_map):
    for key, display in bank_map.items():
        if key in stem:
            print(f"Detected bank: '{display}' (matched '{key}')")
            return display
    raise RuntimeError(f"Cannot detect bank from filename: {stem!r}")

def load_and_filter_db(db_path, sheet, bank_display):
    # read .xls via pandas + xlrd
    df = pd.read_excel(db_path, sheet_name=sheet, engine="xlrd", header=None)
    df.columns = list("ABCDEFGHIJKLMNOPQRSTUVWXYZ")[:df.shape[1]]
    # filter on Column B containing the bank_display text
    filtered = df[df["B"].astype(str).str.contains(bank_display)]
    print(f"Filtered DB to {len(filtered)} rows for '{bank_display}'")
    return filtered





def write_output(matches, out_path: Path, post_date: str):
    """
    Append matched rows into the per-day workbook, skipping duplicates.

    Duplicate key = (E posting_date, U cust_id, S amount).
    Uses counts so legit repeats are allowed beyond what's already in the sheet.
    """
    import openpyxl

    wb = openpyxl.load_workbook(out_path)
    ws = wb["Sheet1"]

    # ---- collect existing counts from already written DZ rows ----
    existing_counts = defaultdict(int)
    for r in range(5, ws.max_row + 1, 2):  # first row of each 2-row block
        e_val = ws.cell(r, 5).value   # E = posting date
        u_val = ws.cell(r, 21).value  # U = customer id
        s_val = ws.cell(r, 19).value  # S = amount
        if e_val is None and u_val is None and s_val is None:
            continue
        try:
            s_float = float(str(s_val).replace(",", "")) if s_val is not None else 0.0
        except ValueError:
            s_float = 0.0
        key = (str(e_val), str(u_val), s_float)
        existing_counts[key] += 1

    # track how many of each key we see in THIS run
    seen_now = defaultdict(int)

    def row_is_empty(r: int) -> bool:
        # Key columns used in a DZ row (1-based): B,C,D,E,F,G,I,J,O,S,U,V
        key_cols = [2,3,4,5,6,7,9,10,15,19,21,22]
        return all(ws.cell(r, c).value is None for c in key_cols)

    # find first empty 2-row block starting at row 5
    row = 5
    while row <= ws.max_row + 1 and not row_is_empty(row):
        row += 2

    ymd    = post_date
    y_str  = ymd[:4]
    m_str  = ymd[4:6]
    md_str = f"{ymd[4:6]}.{ymd[6:]}"

    def fill(r, data, red_cols=None):
        red_cols = red_cols or []
        for col, val in data.items():
            cell = ws[f"{col}{r}"]
            cell.value = val
            if col in red_cols:
                cell.font = RED_FONT
            if col == "S":
                cell.number_format = "#,##0.00"

    written = 0
    for raw_txt, amt, db_row in matches:
        try:
            amt_float = float(str(amt).replace(",", "")) if amt is not None else 0.0
        except ValueError:
            amt_float = 0.0

        cust_id  = db_row["F"]
        clean_nm = db_row["G"]
        hkont    = db_row["C"]
        extra_H  = db_row["H"]
        extra_I  = db_row["I"]

        text_I = f"{md_str} {clean_nm} 暫收款"

        key = (ymd, str(cust_id), amt_float)
        seen_now[key] += 1

        # Skip until we exceed what’s already written earlier today
        if seen_now[key] <= existing_counts.get(key, 0):
            continue

        # Row 1 (DZ)
        r1 = {
            "B": "1000", "C": y_str, "D": "DZ",
            "E": ymd,    "F": ymd,   "G": m_str,
            "I": text_I,
            "J": "NTD",  "O": hkont, "S": amt_float,
            "U": cust_id, "V": text_I,
            "AP": extra_H, "AU": extra_I,
        }
        fill(row, r1, red_cols=["E", "F", "G", "S"])

        # Row 2 (N=5)
        r2 = {
            "L": cust_id,
            "N": "5",
            "S": -amt_float,
            "U": cust_id,
            "V": text_I,
        }
        fill(row + 1, r2)

        row += 2
        written += 2

    wb.save(out_path)
    print(f"Wrote {written} rows into {out_path.name}")

def main():
    args      = parse_args()
    post_date = args.date or datetime.today().strftime("%Y%m%d")

    bank_path = Path(args.file).expanduser()
    parser    = make_parser(bank_path)
    entries   = parser.extract_rows()
    print(f"Loaded {len(entries)} entries from {bank_path.name}")


    # # 1) pick the right parser & extract
    # parser  = make_parser(BANK_FILE)
    # entries = parser.extract_rows()
    # print(f"Loaded {len(entries)} entries from {BANK_FILE.name}")

    # # 2) Detect which bank we’re processing
    # stem = Path(BANK_FILE).stem
    bank_display = detect_bank(bank_path.stem, BANK_MAP)

    # 3) Load & filter the customer DB
    db = load_and_filter_db(DB_FILE, DB_SHEET, bank_display)

    # # 4) Match
    # matches = match_entries_debug(entries, db, FUZZY_THRESHOLD)
    matches, skipped = match_entries_interactive(entries, db, FUZZY_THRESHOLD)
    if skipped:
        log_skipped(skipped, filepath="skipped.csv")

    print(f"DEBUG  → matches found: {len(matches)}")

    out_path = daily_output_path(post_date)
    if not out_path.exists():
        shutil.copy(TEMPLATE_FILE, out_path)  # first bank of the day creates the file 
    # 5) (next: write out your two‐row blocks into the output template)
    write_output(matches, out_path, post_date)

if __name__ == "__main__":
    main()
