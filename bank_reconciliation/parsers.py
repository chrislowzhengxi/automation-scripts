import openpyxl
from pathlib import Path
import pandas as pd
from typing import Union



class BankParserBase:
    def __init__(self, path):        # path = pathlib.Path
        self.path = path

    def extract_rows(self):
        """Return list of (raw_customer_text, amount)."""
        raise NotImplementedError


class CitiParser(BankParserBase):    # your existing logic
    SHEET_NAME     = "Sheet2"
    CUSTOMER_COL   = "E"
    AMOUNT_COL     = "G"
    HEADER_KEYWORD = "細節描述"

    def extract_rows(self):
        # wb = openpyxl.load_workbook(self.path, data_only=True)
        # ws = wb[self.SHEET_NAME]
        ws = load_sheet(self.path, sheet=self.SHEET_NAME)

        # 1) find header row(s)
        hits = [
            r for r in range(1, ws.max_row+1)
            if ws[f"{self.CUSTOMER_COL}{r}"].value == self.HEADER_KEYWORD
        ]
        if not hits:
            raise RuntimeError(f"No '{self.HEADER_KEYWORD}' found in {self.path.name}")

        hdr = hits[1] if len(hits) > 1 else hits[0]
        start = hdr + 2

        # 2) read until blank
        rows = []
        r = start
        while True:
            cust = ws[f"{self.CUSTOMER_COL}{r}"].value
            if cust is None or not str(cust).strip():
                break
            amt = ws[f"{self.AMOUNT_COL}{r}"].value
            rows.append((str(cust).strip(), amt))
            r += 1

        return rows


class CTBCParser(BankParserBase):
    SHEET_NAME     = None       # .xls has only one sheet
    CUSTOMER_COL   = "J"
    AMOUNT_COL     = "E"
    HEADER_KEYWORD = "備註"

    def extract_rows(self):
        # 1) read entire sheet into a DataFrame (no header row)
        # df = pd.read_excel(
        #     self.path,
        #     sheet_name=0,
        #     header=None,
        #     engine="xlrd",
        #     dtype=str   # read everything as strings to preserve formatting
        # )
        df = load_sheet(self.path, sheet=0, header=None)

        # 2) locate your header row by scanning column J
        #    column J → DataFrame column index 9 (0-based)
        header_rows = df[df[9] == self.HEADER_KEYWORD].index
        if header_rows.empty:
            raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")
        hdr = header_rows[0]

        # 3) pull out everything below that until you hit a blank in J
        rows = []
        for idx in range(hdr + 1, len(df)):
            cust = df.at[idx, 9]   # column J
            if pd.isna(cust) or not str(cust).strip():
                break
            amt = df.at[idx, 4]    # column E → index 4
            # convert amount to float if you like
            amt = float(amt.replace(",", "")) if isinstance(amt, str) else amt
            rows.append((cust.strip(), amt))

        print(f"Loaded {len(rows)} entries from {self.path.name}")
        return rows


class MegaParser(BankParserBase):
    """
    Parses 1000-兆豐-*.xlsx
    - Header row has '存入金額' in column F
    - Customer name sits under '備註' in column H
    - Stop reading once column D contains '總計'
    """
    SHEET_NAME     = None        # single‐sheet workbooks
    AMOUNT_COL     = "F"
    CUSTOMER_COL   = "H"
    STOP_COL       = "D"
    HEADER_KEYWORD = "存入金額"
    STOP_TOKEN     = "總計"

    def extract_rows(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        ws = wb.active

        # 1) find header row in column F
        hdr = None
        for r in range(1, ws.max_row + 1):
            if ws[f"{self.AMOUNT_COL}{r}"].value == self.HEADER_KEYWORD:
                hdr = r
                break
        if hdr is None:
            raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")

        # 2) read data until we hit '總計' in column D
        rows = []
        r = hdr + 1
        while r <= ws.max_row:
            # if this row is the grand‐total/subtotal row, stop completely
            if ws[f"{self.STOP_COL}{r}"].value == self.STOP_TOKEN:
                break

            cust = ws[f"{self.CUSTOMER_COL}{r}"].value
            amt  = ws[f"{self.AMOUNT_COL}{r}"].value

            # if customer cell is empty (unlikely for this bank), also stop
            if cust is None or not str(cust).strip():
                break

            rows.append((str(cust).strip(), amt))
            r += 1

        return rows
