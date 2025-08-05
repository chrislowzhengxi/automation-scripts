import openpyxl
from pathlib import Path
import pandas as pd
from utils import load_sheet
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
    Parses 1000-兆豐-*.xls[x]
    - Header row has '存入金額' in column F
    - Customer name sits under '備註' in column H
    - Stop reading once column D contains '總計'
    """
    SHEET      = 0           # single‐sheet workbooks
    AMOUNT_COL = "F"
    CUSTOMER_COL = "H"
    STOP_COL   = "D"
    HEADER_KEYWORD = "存入金額"
    STOP_TOKEN = "總計"

    def extract_rows(self) -> list[tuple[str, float]]:
        # load_sheet will return either:
        #  - openpyxl.Worksheet for .xlsx
        #  - pandas.DataFrame   for .xls
        sheet = load_sheet(self.path, sheet=self.SHEET, header=None)

        # --- XLSX path (openpyxl) ---
        if isinstance(sheet, openpyxl.worksheet.worksheet.Worksheet):
            ws = sheet
            # 1) find header row in column F
            hdr = None
            for r in range(1, ws.max_row + 1):
                if ws[f"{self.AMOUNT_COL}{r}"].value == self.HEADER_KEYWORD:
                    hdr = r
                    break
            if hdr is None:
                raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")

            # 2) walk until STOP_TOKEN in column D
            rows = []
            r = hdr + 1
            while r <= ws.max_row:
                if ws[f"{self.STOP_COL}{r}"].value == self.STOP_TOKEN:
                    break

                cust = ws[f"{self.CUSTOMER_COL}{r}"].value
                amt  = ws[f"{self.AMOUNT_COL}{r}"].value

                if cust is None or not str(cust).strip():
                    break

                rows.append((str(cust).strip(), amt))
                r += 1

            return rows

        # --- XLS path (pandas) ---
        elif isinstance(sheet, pd.DataFrame):
            df: pd.DataFrame = sheet
            # column letters to 0-based indices: D=3, F=5, H=7
            COL_D, COL_F, COL_H = 3, 5, 7

            # 1) find header row by matching HEADER_KEYWORD in column F
            mask = df[COL_F] == self.HEADER_KEYWORD
            if not mask.any():
                raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")
            hdr = mask.idxmax()  # first occurrence

            # 2) iterate until STOP_TOKEN appears in column D
            rows: list[tuple[str, float]] = []
            for idx in range(hdr + 1, len(df)):
                if df.at[idx, COL_D] == self.STOP_TOKEN:
                    break

                cust = df.at[idx, COL_H]
                if pd.isna(cust) or not str(cust).strip():
                    break

                amt = df.at[idx, COL_F]
                # optionally convert comma‐thousands to float
                amt = float(amt.replace(",", "")) if isinstance(amt, str) else amt

                rows.append((cust.strip(), amt))

            return rows

        else:
            raise TypeError(f"Unexpected sheet type: {type(sheet)}")
        

class FubonParser(BankParserBase):
    """
    Parses 1000-富邦-*.xls/.xlsx
    - Header row has '存入金額' in column F
    - Customer name sits under '附言' in column I
    - Stop reading once column A contains '小計'
    """
    SHEET_NAME     = "報表"      
    AMOUNT_COL     = "F"
    CUSTOMER_COL   = "I"
    HEADER_KEYWORD = "存入金額"
    STOP_TOKEN     = "小計"

    def extract_rows(self):
        # load_sheet will return openpyxl.Worksheet for .xlsx, DataFrame for .xls
        sheet = load_sheet(self.path, sheet=self.SHEET_NAME, header=None)

        rows = []
        if isinstance(sheet, openpyxl.worksheet.worksheet.Worksheet):
            # ── .xlsx path
            ws = sheet
            # 1) find header row
            hdr = None
            for r in range(1, ws.max_row + 1):
                if ws[f"{self.AMOUNT_COL}{r}"].value == self.HEADER_KEYWORD:
                    hdr = r
                    break
            if hdr is None:
                raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")

            # 2) walk down until you see '小計' in column A
            r = hdr + 1
            while r <= ws.max_row:
                if ws[f"A{r}"].value == self.STOP_TOKEN:
                    break

                raw = ws[f"{self.AMOUNT_COL}{r}"].value
                cust = ws[f"{self.CUSTOMER_COL}{r}"].value

                if cust is None or not str(cust).strip():
                    break

                # — normalize amt to float if it's a string —
                if isinstance(raw, str):
                    amt = float(raw.replace(",", ""))
                else:
                    amt = raw
                
                rows.append((str(cust).strip(), amt))
                r += 1

        else:
            # ── .xls path via pandas DataFrame
            df: pd.DataFrame = sheet
            # column letters → zero-based indices: A=0, F=5, I=8
            # 1) find header row where col 5 == HEADER_KEYWORD
            header_rows = df[df[5] == self.HEADER_KEYWORD].index
            if header_rows.empty:
                raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")
            hdr = header_rows[0]

            # 2) read until you hit STOP_TOKEN in column 0
            for idx in range(hdr + 1, len(df)):
                if df.at[idx, 0] == self.STOP_TOKEN:
                    break

                cust = df.at[idx, 8]  # 附言
                amt  = df.at[idx, 5]  # 存入金額

                if pd.isna(cust) or not str(cust).strip():
                    break

                # optional: convert amt string with commas to float
                if isinstance(amt, str):
                    amt = float(amt.replace(",", ""))

                rows.append((str(cust).strip(), amt))

        print(f"🔍 Loaded {len(rows)} entries from {self.path.name}")
        return rows
    
    
class SinopacParser(BankParserBase):
    """
    Parses 1000-永豐-*.xls/.xlsx
    - Header row has '存入' in column F
    - Customer name sits under '備註' in column J
    - Stop when you hit a truly blank customer cell
    """
    SHEET_NAME     = "工作表1"   # or leave None to use the first sheet
    AMOUNT_COL     = "F"
    CUSTOMER_COL   = "J"
    HEADER_KEYWORD = "存入"

    def extract_rows(self):
        sheet = load_sheet(self.path,
                           sheet=self.SHEET_NAME,
                           header=None)

        rows = []
        # .xlsx path
        if isinstance(sheet, openpyxl.worksheet.worksheet.Worksheet):
            ws = sheet
            # 1) locate header row in column F
            hdr = None
            for r in range(1, ws.max_row+1):
                if ws[f"{self.AMOUNT_COL}{r}"].value == self.HEADER_KEYWORD:
                    hdr = r
                    break
            if hdr is None:
                raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")

            # 2) read down until customer cell is blank
            r = hdr + 1
            while r <= ws.max_row:
                cust = ws[f"{self.CUSTOMER_COL}{r}"].value
                amt  = ws[f"{self.AMOUNT_COL}{r}"].value

                # stop on blank customer
                if cust is None or not str(cust).strip():
                    break

                cust_text = str(cust).strip()
                # normalize amt → float if it's text
                if isinstance(amt, str):
                    amt = float(amt.replace(",", ""))
                rows.append((cust_text, amt))
                r += 1

        # .xls path via pandas
        else:
            df: pd.DataFrame = sheet
            # A=0, F=5, J=9
            header_rows = df[df[5] == self.HEADER_KEYWORD].index
            if header_rows.empty:
                raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")
            hdr = header_rows[0]

            for idx in range(hdr + 1, len(df)):
                cust = df.at[idx, 9]
                if pd.isna(cust) or not str(cust).strip():
                    break
                amt = df.at[idx, 5]
                if isinstance(amt, str):
                    amt = float(amt.replace(",", ""))
                rows.append((str(cust).strip(), amt))

        print(f"🔍 Loaded {len(rows)} entries from {self.path.name}")
        return rows
