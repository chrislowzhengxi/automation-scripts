
# README.md

## English

### 1. Project Overview

This project automates processing of **SAP 科餘 Excel exports** into a structured workbook with grouped account sheets and a consolidated **說明 (Explanations)** sheet.

It is designed for Treasury/Accounting workflows, where multiple bank/SAP exports need to be merged, grouped by G/L account, and reviewed against various compliance checks (e.g., >30 days, >90 days, supplier overlaps).

Typical workflow:

1. Merge raw export files.
2. Group by G/L account into separate sheets.
3. Review the **說明 sheet**, which contains compliance checks, summaries, and cross-account comparisons.

---

### 2. Key Features

* Merge multiple SAP export Excel files into one workbook.
* Automatically group transactions by **G/L科目** (account number).
* Apply consistent styling, freeze panes, and column grouping.
* **說明 sheet** enhancements:

  * Partial column grouping (D–E, G, K–L, U–W; A–B remain visible).
  * Extra 「說明」 column appended after every data block.
  * Column R widened for longer texts.
* Date parsing and automatic Excel date formatting (`m/d/yyyy`).
* Highlighting: matched rows are marked in the grouped account sheets.
* **Section 9 cross-checks**:

  * Detect overlapping supplier numbers between

    * `12580100` and `21780101`
    * `12810100` and `22280201`
  * Output common suppliers list and filtered detail tables for each account.

---

### 3. Requirements

* **Python**: 3.9 or higher
* **Dependencies** (from `requirements.txt`):

  ```txt
  pandas>=2.0.0
  openpyxl>=3.1.0
  lxml>=4.0.0
  html5lib>=1.0.0
  xlrd>=2.0.0
  ```

Install with:

```bash
pip install -r requirements.txt
```

---

### 4. Input Files

1. **Export workbook**

   * Must contain G/L科目, 文件號碼, 過帳日期, 結清文件, and supplier columns.
   * File is usually SAP 科餘 export.
2. **Mapping workbook**

   * Example: `會計科目對照表.xlsx`
   * Maps account codes (G/L numbers) to names.
3. (Optional) **Reference file**

   * Used for consistent column order when merging.

---


### 5. Usage

#### A. Launch the Merger GUI

1. Double-click **`Open Merger GUI.bat`**.
2. The GUI (see screenshot) will appear.

* **Reference file (recommended):** Select your 科餘 export reference file.
* **Input Excel files to merge:** Choose one or more SAP export files.
* **Output file name:** Type the desired combined file name (e.g., `combined.xlsx`).
* **Cutoff date (optional):** Enter `YYYY-MM-DD` or leave blank (defaults to today).
* Optionally tick **“Keep duplicate rows”** if you don’t want duplicates removed.

Click **Merge** to generate the merged Excel file.

#### B. Group by G/L Accounts

After merging, run the `group_by_gl.py` script on the merged file to generate grouped sheets and the 說明 sheet.
This step will apply the enhancements (column grouping, 說明 column, and the cross-checks in point 9).


---

### 6. Output Behavior

**Account Sheets**

* One sheet per G/L科目.
* Columns B–X copied with styles.
* Freeze header row.
* Filters enabled.
* Column grouping: A–B, D–E, G, K–L, U–W.

**說明 Sheet**

* Contains all compliance questions (1–9).
* Column grouping (D–E, G, K–L, U–W).
* No filter on row 1, only frozen panes.
* 「說明」 column appended after each data block.
* Column R widened for text.

**Cross-Check Logic (Section 9)**

* If supplier numbers overlap between the specified pairs, the sheet shows:

  * A list of common suppliers.
  * Detail rows from both accounts.

---

### 7. Troubleshooting

* **Columns missing** → Check your export contains the required headers.
* **Sheet name collisions** → Automatically handled with unique names.
* **Excel shows `#####`** → Column too narrow; expand manually if needed.
* **Grouping not working** → Ensure you’re opening in Excel (some viewers ignore outline symbols).

---

### 8. Contribution Guidelines

* Open issues for bugs or enhancements.
* PRs should include:

  * Clear description of changes.
  * Before/after screenshots if layout or formatting changes are involved.

---

### 9. License

This project is intended for internal Treasury/Accounting automation. Add license info if needed.

---

## 繁體中文

### 1. 專案簡介

本專案用於自動化處理 **SAP 科餘 Excel 匯出檔**，將其轉換為結構化工作簿：

* 依 **G/L 科目** 分組成多個分頁
* 產出整合的 **「說明」sheet**，內含各種檢查項目與比對結果

典型流程：

1. 合併多個匯出檔案
2. 依科目分組
3. 檢閱 **說明 sheet** 的各項檢查與比對結果

---

### 2. 主要功能

* 合併多個 SAP 科餘匯出檔
* 自動依 **G/L科目** 分頁
* 樣式、凍結窗格、欄位群組自動化
* **說明 sheet 增強**：

  * 部分欄位群組（D–E、G、K–L、U–W；A–B 保留顯示）
  * 新增「說明」欄位（每個表格後皆會新增）
  * R 欄位加寬以容納文字
* 自動日期解析與 Excel 格式化
* 在科目分頁中高亮對應列
* **第 9 點交叉比對**：

  * `12580100` ↔ `21780101`
  * `12810100` ↔ `22280201`
  * 顯示共同供應商清單與各科目明細

---

### 3. 系統需求

* **Python**: 3.9 以上
* **相依套件** (`requirements.txt`):

  ```txt
  pandas>=2.0.0
  openpyxl>=3.1.0
  lxml>=4.0.0
  html5lib>=1.0.0
  xlrd>=2.0.0
  ```

安裝套件：

```bash
pip install -r requirements.txt
```

---

### 4. 輸入檔案

1. **匯出報表**

   * 必須包含：G/L科目、文件號碼、過帳日期、結清文件、供應商等欄位
2. **對照表**

   * 範例：`會計科目對照表.xlsx`
   * 用於將科目代碼對應名稱
3. （選用）**參考檔**

   * 用於保持合併後欄位順序一致



---

### 5. 使用方式

#### A. 開啟合併工具 GUI

1. 直接雙擊 **`Open Merger GUI.bat`**。
2. GUI 介面將會出現（如截圖）。

* **Reference file (建議):** 選擇科餘匯出參考檔案。
* **Input Excel files to merge:** 選擇一個或多個 SAP 匯出檔案。
* **Output file name:** 輸入合併後的檔名 (如 `combined.xlsx`)。
* **Cutoff date (選填):** 輸入 `YYYY-MM-DD`，或留白 (預設為今天)。
* 可勾選 **「Keep duplicate rows」** 以保留重複列。

按下 **Merge** 即可產生合併後的 Excel 檔。

#### B. 按 G/L 科目分組

合併完成後，執行 `group_by_gl.py`，指定合併後檔案，即可自動產生分組表與「說明」sheet。
此步驟會套用增強功能（欄位群組、「說明」欄，以及第 9 點交叉檢查）。

---


### 6. 輸出與行為

**科目分頁**

* 每個 G/L科目一個分頁
* 保留欄位 B–X 與原始樣式
* 凍結第一列
* 自動篩選啟用
* 欄位群組：A–B、D–E、G、K–L、U–W

**說明 sheet**

* 包含問題 1–9
* 欄位群組（D–E、G、K–L、U–W）
* 第一列無篩選（僅凍結）
* 每個表格新增「說明」欄
* R 欄加寬

**第 9 點交叉比對**

* 若供應商號碼重複，會列出共同清單並顯示各科目明細

---

### 7. 疑難排解

* **欄位缺漏** → 請確認匯出報表包含必要欄位
* **分頁名稱衝突** → 程式會自動處理
* **Excel 顯示 `#####`** → 欄位寬度不足，手動放大即可
* **群組無法摺疊** → 確認在 Excel 中開啟（部分檢視器不支援）

---

### 8. 開發參與

* 回報問題或提出需求請開 Issue
* PR 請包含：

  * 清楚的變更描述
  * 若涉及版面/格式，請附上前後對照圖

---


