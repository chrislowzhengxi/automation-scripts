import pandas as pd
import openpyxl
from rapidfuzz import process, fuzz
from pathlib import Path

BASE_DIR = Path("~/Downloads/Banks").expanduser()

BANK_FILE       = BASE_DIR / "花旗銀行對帳單-20250625.xlsx"

bank_key = Path(BANK_FILE).stem.split("對帳單")[0]
print(bank_key)