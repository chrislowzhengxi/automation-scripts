import openpyxl
import pandas as pd
from pathlib import Path
from typing import Union

def load_sheet(path: Union[str, Path],
               sheet: Union[int, str] = 0,
               header: Union[int, None] = None) -> Union[openpyxl.worksheet.worksheet.Worksheet, pd.DataFrame]:
    """
    Load a sheet from .xlsx (via openpyxl) or .xls (via pandas/xlrd).
    
    - path: path to your bank file
    - sheet: sheet name or index (0-based for pandas, name or index for openpyxl)
    - header: only for pandas read_excel; which row to treat as header (None = all rows are data)
    
    Returns:
      - openpyxl Worksheet (if .xlsx)
      - pd.DataFrame (if .xls)
    """
    p = Path(path)
    ext = p.suffix.lower()
    
    if ext == ".xlsx":
        wb = openpyxl.load_workbook(p, data_only=True)
        # if they passed an integer, pick by index; otherwise by name
        if isinstance(sheet, int):
            return wb.worksheets[sheet]
        else:
            return wb[sheet]
    
    elif ext == ".xls":
        # pandas with xlrd
        return pd.read_excel(
            p,
            sheet_name=sheet,
            header=header,
            engine="xlrd",
            dtype=str  # read everything as string so you can strip/convert
        )
    else:
        raise ValueError(f"Unsupported extension {ext!r}, expected .xls or .xlsx")
