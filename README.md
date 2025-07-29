# Wistron Finance Automation

This repository contains internal automation scripts developed for the Wistron finance team to streamline repetitive reporting and reconciliation tasks. The project is organized into submodules for revenue reporting and bank reconciliation.


## 🔧 Scripts Overview

### `revenue_update/update_revenue.py`

Automates the monthly revenue update process:

- Pulls cumulative revenue from `RPTIS10_I_A01_YYYYMM.xlsx`.
- Locates the correct `High-Monthy-營收公告-11404.xlsx` Excel file or falls back to the latest one.
- Inserts the new revenue, calculates net revenue, and updates formulas.
- Applies formatting and hides older columns to keep the view clean.

#### Power Automate Integration

This script is triggered by Power Automate Desktop, which passes in a `YearMonth` argument (e.g., `202504`). The automation flow:

1. Extracts the relevant Excel files from a shared location or download folder.
2. Runs this Python script via a command line action.
3. Optionally renames or moves the result after processing.

See `revenue_update/revenue_power_automate.txt` for a copy of the Power Automate flow logic.

#### Usage (via Power Automate):
```
python update_revenue.py 202504
```

### `bank_reconciliation/bank3.py`

Processes bank statement files (e.g., from Citibank):

- Extracts transaction details (customer names and amounts) from `.xlsx` bank reports.
- Performs fuzzy matching of customer names against a master customer list.
- Copies the matching entries (2-row blocks) into the accounting voucher template (`會計憑證導入模板 - 空白檔案.xlsx`).
- Updates each matched entry’s transaction date and payment amount.
- Designed to match names even if they are only partially or approximately correct.

> Example input: `花旗銀行對帳單-20250625.xlsx`  
> Example customer master: `會計憑證導入模板 - 1000 客戶資料檔.xlsx`

⚠ This script expects the files to be in the `Downloads` folder and only modifies the output template — never the customer master.

#### Key features:
- Two-row template blocks for each customer
- Red font and formatting applied to amount fields
- Can be extended to support other banks (e.g., HSBC, SCB, etc.)

### 📌 Requirements
Python 3.9+

Install dependencies:
```
pip install -r requirements.txt
```
