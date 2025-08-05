import openpyxl
from pathlib import Path


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
        wb = openpyxl.load_workbook(self.path, data_only=True)
        ws = wb[self.SHEET_NAME]

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
    """
    Parses 1000-中信-*.xls
    - Header row has '備註' (customer) in column J
    - Credit amount lies in column E under '轉入/匯款金額'
    """
    SHEET_NAME     = None      # CTBC is the only sheet
    CUSTOMER_COL   = "J"
    AMOUNT_COL     = "E"
    HEADER_KEYWORD = "備註"

    def extract_rows(self):
        wb = openpyxl.load_workbook(self.path, data_only=True)
        ws = wb.active

        # find the “備註” header
        hdr = None
        for r in range(1, ws.max_row+1):
            if ws[f"{self.CUSTOMER_COL}{r}"].value == self.HEADER_KEYWORD:
                hdr = r
                break
        if hdr is None:
            raise RuntimeError(f"No '{self.HEADER_KEYWORD}' in {self.path.name}")

        rows = []
        r = hdr + 1
        while True:
            cust = ws[f"{self.CUSTOMER_COL}{r}"].value
            if cust is None or not str(cust).strip():
                break
            amt = ws[f"{self.AMOUNT_COL}{r}"].value
            rows.append((str(cust).strip(), amt))
            r += 1

        return rows
