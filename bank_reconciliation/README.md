## 📁 Sample Excel Files

This repository includes sanitized Excel files for safe testing of the bank reconciliation automation:

- `bank_reconciliation/sanitized_bank.xlsx`  
- `bank_reconciliation/sanitized_master.xls`

These files are cleaned versions of the original templates used in production. All sensitive customer names and identifiers have been replaced with generic placeholders (e.g., `Customer 1`, `Customer 2`, etc.), while the structure and format are preserved to ensure the automation scripts work as intended.

**⚠️ Do not edit these files directly.**  
To run the automation scripts, copy them to your local `Downloads` folder, as the scripts are designed to read input files from there.

If you're working within the company network and using live data, make sure you do **not** push unsanitized files to this repository.
