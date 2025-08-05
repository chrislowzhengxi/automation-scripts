#!/usr/bin/env python3
from pathlib import Path
from parsers import CitiParser, CTBCParser, MegaParser, FubonParser, SinopacParser, ESunParser, BankParserBase

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
OUTPUT_FILE = BASE_DIR / "會計憑證導入模板 - 空白檔案.xlsx"
RED_FONT    = Font(color="FF0000")

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

# ─────────────── 4) MATCH & DEBUG ───────────────
def match_entries_debug(entries, db, threshold=80):
    """Return [(raw_text, amount, db_row)] with verbose logs."""
    keywords = db["E"].astype(str).str.strip().tolist()
    matches  = []

    for raw_txt, amt in entries:
        print("\n🔎 BANK ROW")
        print(f"   Text   : {raw_txt!r}")
        print(f"   Amount : {amt}")

        # 4-a) exact substring in db["E"]
        clean   = raw_txt.replace(" ", "")
        subset  = db[db["E"].apply(lambda k: str(k).replace(' ', '') in clean)]
        if not subset.empty:
            hit = subset.iloc[0]
            print("   ✅ Exact match:")
            print(f"      Keyword     : {hit['E']!r}")
            print(f"      Customer ID : {hit['F']}  Clean Name : {hit['G']!r}")
            matches.append((raw_txt, amt, hit))
            continue

        # 4-b) fuzzy fallback
        best = process.extractOne(
            clean, keywords, scorer=fuzz.partial_ratio
        )
        if best:
            best_kw, score, _ = best
            print(f"   ➡️  Fuzzy best : {best_kw!r}  (score {score:.1f})")
            if score >= threshold:
                idx = keywords.index(best_kw)
                hit = db.iloc[idx]
                print("   ✅ Accepted fuzzy match")
                matches.append((raw_txt, amt, hit))
                continue
            else:
                print(f"   ⚠️  Score {score:.1f} < threshold {threshold}")
        else:
            print("   ⚠️  No fuzzy candidate at all")

    print(f"\n🔗 Matched {len(matches)}/{len(entries)} rows")
    return matches

# ─────────────── 5) WRITE OUTPUT ───────────────
def write_output(matches, template_path, post_date):
    """post_date = 'YYYYMMDD' string supplied by --date"""
    wb = openpyxl.load_workbook(template_path)
    ws = wb["Sheet1"]                 # adjust if the tab is named differently

    # ── 1) Find the first completely empty row in column A, beginning at row 5
    for row in range(5, ws.max_row+2):
        if ws.cell(row, 1).value is None:
            break

    # ── 2) date helpers ───────────────────────────────────────────────
    ymd    = post_date                         # "20250625"
    y_str  = ymd[:4]                           # "2025"
    m_str  = ymd[4:6]                          # "06"
    md_str = f"{ymd[4:6]}.{ymd[6:]}"           # "06.25"

    def fill(r, data, red_cols=None):
        red_cols = red_cols or []
        for col, val in data.items():
            cell = ws[f"{col}{r}"]
            cell.value = val
            if col in red_cols:
                cell.font = RED_FONT
            # — force two decimal places on column S —
            if col == "S":
                cell.number_format = "#,##0.00"

    for raw_txt, amt, db_row in matches:
        cust_id  = db_row["F"]
        clean_nm = db_row["G"]
        hkont    = db_row["C"]          # adjust if HKONT is another column
        extra_H  = db_row["H"]
        extra_I  = db_row["I"]

        # ── Row 1 (DZ row) ───────────────────────────────────────────
        r1 = {
            "B": "1000", "C": y_str, "D": "DZ",
            "E": ymd,    "F": ymd,   "G": m_str,
            "I": f"{md_str} {clean_nm} 暫收款",
            "J": "NTD",  "O": hkont, "S": amt,
            "U": cust_id, "V": f"{md_str} {clean_nm} 暫收款",
            "AP": extra_H, "AU": extra_I,
        }
        # Row-1: mark E F G red
        fill(row, r1, red_cols=["E", "F", "G", "S"])

        # ── Row 2 (ID / N=5 row) ─────────────────────────────────────
        r2 = {
            "L": cust_id,
            "N": "5",
            "S": -amt,
            "U": cust_id,
            "V": f"{md_str} {clean_nm} 暫收款",
        }
        fill(row+1, r2)

        row += 2                     # Jump two rows 

    # ── 3) sort by HKONT in two-row blocks ───────────────────────────
    data = list(ws.iter_rows(min_row=5, values_only=True))
    blocks = [data[i:i+2] for i in range(0, len(data), 2)]   # 2 rows + blank
    blocks.sort(key=lambda blk: str(blk[0][14] or ""))       # col O index 14
    # rewrite
    ws.delete_rows(5, ws.max_row)
    r = 5
    for blk in blocks:
        for vals in blk:
            for c, v in enumerate(vals, start=1):
                ws.cell(row=r, column=c, value=v)
            r += 1

    # ── 4) re-apply styles to each 2-row customer block ────────────────────
    # Columns:  E=5  F=6  G=7   S=19
    for rr in range(5, ws.max_row + 1, 2):          # rr = first row of each pair
        # 1) paint E/F/G on the first row red
        for cc in (5, 6, 7):
            ws.cell(rr, cc).font = RED_FONT

        # 2) format column S on BOTH rows; red-font only on the positive row
        for r, make_red in ((rr, True), (rr + 1, False)):
            s_cell = ws.cell(r, 19)                 # column S
            s_cell.number_format = "#,##0.00"
            if make_red:                            # row rr (positive amount)
                s_cell.font = RED_FONT

            
    wb.save(template_path)
    print(f"✅ Wrote {len(matches)*2} rows into {template_path.name}")



def main():
    args      = parse_args()
    post_date = args.date or datetime.today().strftime("%Y%m%d")

    bank_path = Path(args.file).expanduser()
    parser    = make_parser(bank_path)
    entries   = parser.extract_rows()
    print(f"🔍 Loaded {len(entries)} entries from {bank_path.name}")


    # # 1) pick the right parser & extract
    # parser  = make_parser(BANK_FILE)
    # entries = parser.extract_rows()
    # print(f"🔍 Loaded {len(entries)} entries from {BANK_FILE.name}")

    # # 2) Detect which bank we’re processing
    # stem = Path(BANK_FILE).stem
    bank_display = detect_bank(bank_path.stem, BANK_MAP)

    # 3) Load & filter the customer DB
    db = load_and_filter_db(DB_FILE, DB_SHEET, bank_display)

    # # 4) Match
    matches = match_entries_debug(entries, db, FUZZY_THRESHOLD)

    print(f"DEBUG  → matches found: {len(matches)}")
    
    # 5) (next: write out your two‐row blocks into the output template)
    write_output(matches, OUTPUT_FILE, post_date)

if __name__ == "__main__":
    main()
    # args     = parse_args()
    # post_date = args.date or datetime.today().strftime("%Y%m%d")
    # main(post_date)
