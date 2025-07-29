This folder contains the Python script and Power Automate flow for the monthly revenue update automation.

Files:
- update_revenue.py        ← Main script that updates the revenue tracker Excel file
- revenue_power_automate.txt  ← Text copy of the Power Automate Desktop flow (Ctrl+A → Ctrl+C → paste here)

Usage Instructions:
1. Place the required Excel files (e.g., RPTIS10_I_A01_YYYYMM.xlsx and High-Monthy-營收公告-11404.xlsx) in your Downloads folder.
2. Power Automate passes the year-month (e.g., 202504) as an argument to the Python script.
3. The script calculates new revenue, updates formulas, hides old columns, and applies formatting to the monthly Excel tracker.
4. The result is saved back to the same file.

Notes:
- If the target tracker file is missing, the script will attempt to find the most recent one and create a fallback.
- All formatting and calculations are done in-place.
- Script logs will be saved as log.txt in this folder.
