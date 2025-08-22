from pathlib import Path
import argparse
import win32com.client as win32

def replace_ntd_with_usd_in_rates(xls_path: Path, sheet_name: str = "Summary"):
    xls_path = str(xls_path.resolve())
    excel = win32.gencache.EnsureDispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(xls_path)
        sheets = {ws.Name: ws for ws in wb.Worksheets}
        ws = sheets.get(sheet_name, wb.Worksheets(1))

        # find last used row in column B
        xlUp = -4162
        last_row = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row

        row = 1
        while row <= last_row:
            val = ws.Cells(row, 2).Value  # column B
            if isinstance(val, str) and val.strip().upper() == "NTD":
                ws.Cells(row, 2).Value = "USD"
                print(f"Replaced NTD → USD at row {row}")
                break
            row += 1

        wb.Save()
    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        excel.Quit()


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Fix NTD→USD in Ending & Avg workbook")
    ap.add_argument("--period", required=True, help="Month in YYYYMM format, e.g. 202504")
    args = ap.parse_args()

    xls_file = Path(
        fr"C:\Users\TP2507088\Downloads\Automation\ytm_forms\data\template\關係人\{args.period} Ending 及 Avg (資通版本).xls"
    )
    if not xls_file.exists():
        raise FileNotFoundError(f"Rates file not found: {xls_file}")

    replace_ntd_with_usd_in_rates(xls_file)
