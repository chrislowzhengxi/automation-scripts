
# Bank Reconciliation Automation  
# 銀行對帳單自動化工具

### *Scroll to the bottom for the English version 

## 1. 使用者指南 (User Guide)

### 1.1 工具簡介 (Overview)
本工具可自動化處理多家銀行的對帳單，將交易資料比對至會計系統的客戶資料庫，並依每天的入帳日期，輸出符合格式的會計憑證檔案。  
它同時會自動避免重複紀錄，並在同一天的多次執行中建立新的檔案（如 `YYYYMMDD-2.xlsx`, `YYYYMMDD-3.xlsx`）。

---

### 1.2 安裝步驟 (Installation Steps)

#### Windows 使用者 (Windows Users)
1. **下載整個專案資料夾** (`bank_reconciliation`) 到您的電腦（例如 `Downloads` 資料夾）。
2. 確保已安裝 **Python 3.10+**：
   - 打開「命令提示字元」，輸入 `python --version` 檢查。
3. 安裝必要套件（開啟命令提示字元，切換到專案資料夾，輸入）：  
   ```bash
   pip install rapidfuzz pandas openpyxl xlrd

4. 執行圖形介面：

   ```bash
   python run_gui.py
   ```

   或是到 `bank_reconciliation` 中點選 `Run GUI`


---

### 1.3 使用介面說明 (User Interface Guide)

啟動後會看到主視窗，分為數個區域：
After launching, you’ll see the main window with several sections:

#### **上方控制列 (Top Controls)**

* **Posting date (YYYYMMDD)**

  * 預設為今天，可手動輸入其他日期。
  * "Today" 按鈕可快速填入今天日期。
* **Add files…**

  * 選取銀行對帳單 Excel 檔案，可一次選多個。
* **Remove selected** / **Clear**

  * 移除選中的檔案或清空全部檔案列表。

#### **檔案列表 (File List)**

顯示目前待處理的對帳單檔案。

#### **Run 按鈕 (Run button)**

點擊後，依序處理列表中的檔案。

* 工具會：

  1. 自動判斷是哪一家銀行。
  2. 從銀行對帳單中提取交易資訊。
  3. 與會計客戶資料庫比對。
  4. 將符合的資料寫入當日的會計憑證檔案：

     * 第一次執行 → `YYYYMMDD.xlsx`
     * 第二次同日執行 → `YYYYMMDD-2.xlsx`
     * 以此類推。

#### **Open output folder**

打開輸出檔案所在的資料夾（通常是 `Downloads/Banks`）。

#### **Log 區域**

顯示處理過程，包括：

* 偵測到的銀行
* 成功寫入的筆數
* 重複跳過的筆數
* 輸出檔案名稱

---

### 1.4 輸出結果 (Output Results)

* 輸出的 Excel 檔位於 `Downloads/Banks` 資料夾。
* 同一日期的多次執行，會依序建立 `-2`, `-3` 後綴檔案。
* 每個檔案只包含該次執行新增的資料，避免重複。

---

## 2. 開發者指南 (Developer Guide)

### 2.1 檔案結構 (File Structure)

```
bank.py        # 核心邏輯：讀取銀行檔案、比對客戶資料庫、寫入輸出檔案
parsers.py     # 各銀行專用的資料解析類別 (e.g., CitiParser, CTBCParser)
fuzzy_matcher.py # 模糊比對名稱與客戶資料
utils.py       # 共用工具，例如記錄跳過的項目
run_gui.py     # Tkinter 圖形介面啟動入口
```

---

### 2.2 處理流程 (Processing Flow)

1. **GUI (`run_gui.py`)**

   * 收集使用者選擇的檔案與日期
   * 呼叫 `bank.py` 以命令列方式處理每一檔案

2. **銀行檔案解析 (`parsers.py`)**

   * `make_parser()` 根據檔名決定使用哪個 parser 類別
   * 每個 parser 提取：

     * 客戶名稱
     * 金額
     * 交易細節

3. **客戶資料比對 (`fuzzy_matcher.py`)**

   * 根據名稱與資料庫模糊匹配
   * 可人工確認或輸入 ID

4. **輸出 (`bank.py`)**

   * `next_versioned_output_path()` 決定輸出檔案名稱（含 `-2`, `-3` 後綴）
   * `collect_existing_counts()` 聚合當日所有舊檔的紀錄
   * `write_output()` 寫入新檔案，僅保留新的紀錄

---

### 2.3 關鍵邏輯變更 (Key Logic Change)

* 舊版：同一天的多次執行 → 資料追加到同一檔案
* 新版：同一天的多次執行 → 每次新建 `-2`, `-3` 檔案，只包含新增資料
  → 減少混淆並保留批次處理紀錄

---

### 2.4 開發注意事項 (Dev Notes)

* **Template file**: `TEMPLATE_FILE` 指向空白的會計憑證模板
* **Output folder**: 預設為 `~/Downloads/Banks`
* **Duplicate key**: `(E date, U cust_id, S amount)`，跨檔案檢查
* **Adding new banks**:

  * 實作 parser 類別 → `parsers.py`
  * 在 `PARSER_REGISTRY` 中註冊銀行關鍵字與類別
* **Testing**: 測試需包含同日多批次的情境，確認 `-2`、`-3` 檔案正確產生

---


## English Version 

## 1. User Guide

### 1.1 Overview

This tool automates the processing of multiple banks’ statements, matches transaction data against an accounting system’s customer database, and outputs properly formatted voucher files by posting date.
It automatically prevents duplicate records, and for multiple runs on the same day, it creates new files with suffixes such as `YYYYMMDD-2.xlsx`, `YYYYMMDD-3.xlsx`.

---

### 1.2 Installation Steps

#### Windows Users

1. **Download the entire project folder** (`bank_reconciliation`) to your computer (e.g., into the `Downloads` folder).
2. Make sure **Python 3.10+** is installed:

   * Open Command Prompt and type `python --version` to check.
3. Install the required packages (open Command Prompt, navigate to the project folder, and type):

   ```bash
   pip install rapidfuzz pandas openpyxl xlrd
   ```
4. Run the graphical interface:

   ```bash
   python run_gui.py
   ```

   Or, go to the `bank_reconciliation` folder and double-click **Run GUI**.

---

### 1.3 User Interface Guide

After launching, you will see the main window divided into several sections:

#### **Top Controls**

* **Posting date (YYYYMMDD)**

  * Defaults to today; you can enter another date manually.
  * The "Today" button quickly fills in today’s date.
* **Add files…**

  * Select bank statement Excel files; you can select multiple files at once.
* **Remove selected** / **Clear**

  * Remove selected files from the list or clear all files.

#### **File List**

Displays the list of bank statement files to be processed.

#### **Run button**

When clicked, processes each file in the list sequentially:

1. Automatically detects the bank.
2. Extracts transaction data from the bank statement.
3. Matches transactions to the accounting customer database.
4. Writes matched data to the daily voucher file:

   * First run → `YYYYMMDD.xlsx`
   * Second run on the same date → `YYYYMMDD-2.xlsx`
   * And so on.

#### **Open output folder**

Opens the folder containing the output files (usually `Downloads/Banks`).

#### **Log area**

Displays processing details, including:

* Detected bank
* Number of records successfully written
* Number of duplicates skipped
* Name of the output file

---

### 1.4 Output Results

* Output Excel files are located in the `Downloads/Banks` folder.
* Multiple runs on the same date will generate files with suffixes `-2`, `-3`, etc.
* Each file contains only the new data added in that run, avoiding duplicates.

---

## 2. Developer Guide

### 2.1 File Structure

```
bank.py          # Core logic: reads bank files, matches customer database, writes output files
parsers.py       # Bank-specific parser classes (e.g., CitiParser, CTBCParser)
fuzzy_matcher.py # Fuzzy matching between names and customer data
utils.py         # Shared utilities, e.g., logging skipped items
run_gui.py       # Tkinter GUI entry point
```

---

### 2.2 Processing Flow

1. **GUI (`run_gui.py`)**

   * Collects user-selected files and date
   * Calls `bank.py` from the command line to process each file

2. **Bank file parsing (`parsers.py`)**

   * `make_parser()` decides which parser class to use based on filename
   * Each parser extracts:

     * Customer name
     * Amount
     * Transaction details

3. **Customer matching (`fuzzy_matcher.py`)**

   * Matches names to the database using fuzzy matching
   * Allows manual confirmation or ID entry

4. **Output (`bank.py`)**

   * `next_versioned_output_path()` determines the output file name (with `-2`, `-3` suffixes)
   * `collect_existing_counts()` aggregates records from all previous files for the date
   * `write_output()` writes a new file containing only new records

---

### 2.3 Key Logic Change

* Old behavior: multiple runs on the same day → data appended to the same file
* New behavior: multiple runs on the same day → new file `-2`, `-3` created for each run, containing only new records
  → Reduces confusion and keeps batch history intact

---

### 2.4 Dev Notes

* **Template file**: `TEMPLATE_FILE` points to the blank voucher template
* **Output folder**: defaults to `~/Downloads/Banks`
* **Duplicate key**: `(E date, U cust_id, S amount)` is checked across files
* **Adding new banks**:

  * Implement a parser class in `parsers.py`
  * Register the bank keyword and class in `PARSER_REGISTRY`
* **Testing**: include scenarios with multiple runs on the same date to confirm correct `-2`, `-3` file creation

---


