

# 📊 YTM Forms Automation

This tool automates the monthly preparation of **YTM forms** (上櫃公司與關係人間重要交易資訊).
It wraps all the copy/paste, formulas, and external lookups into a single Python script — with an optional GUI for non-technical users.

---

## 🚀 Features

* Copy and align data into **2-3.銷貨明細** and **4-3.應收關係人科餘**.
* Auto-generate month structures in **1-1.公告(元)** (handles snapshots, new period insertion, totals).
* Insert formulas for:

  * Differences (E–N), ratios, and totals.
  * External references to `RPTIS10`, `MRS0034`, `MRS0014`.
  * Exchange rate lookups (匯率 / 換算台幣).
  * Related-party names from `關係企業(人).xls`.
* Output formatted with borders, accounting/percentage styles.
* **GUI mode** so coworkers can run it without touching the command line.
* Remembers last inputs between runs.

---

## 📂 Repository Layout

```
ytm_forms/
  data/
    template/   <- base templates + related-party reference files
    output/     <- generated Excel workbooks
  scripts/
    fill_updated.py   <- core automation script (CLI)
run_gui_fill_updated.py   <- Tkinter GUI wrapper
```

---

## 🛠️ Requirements

* Python 3.9+
* Packages:

  * `openpyxl`
  * `pandas`
  * `xlrd` (for legacy `.xls` files)
* Windows (paths and Excel external refs are Windows-style)

Install dependencies:

```bash
pip install -r requirements.txt
```

---

## ⚡ Command Line Usage

Example:

```bash
python -m ytm_forms.scripts.fill_updated \
  --template "ytm_forms/data/template/Template 202504YTM.xlsx" \
  --task all \
  --period 202504 \
  --announce-sheet "1-1.公告(元)"
```

**Arguments:**

* `--template` : Path to the template workbook.
* `--task` : Which part to run (`copy_4_3`, `copy_2_3`, `both`, `announce_structure`, `all`).
* `--period` : Year/month, format `YYYYMM`.
* `--announce-sheet` : Announce sheet name (default `1-1.公告(元)`).
* `--inplace` : Overwrite the template in place.
* `--out` : Custom output path (if not in-place).
* Optional overrides: `--src-43`, `--src-23`, `--rates-path`, `--relparty-path`, `--rptis10-path`, `--mrs0034-path`, `--mrs0014-path`.

Outputs are saved under:

```
ytm_forms/data/output/copy_<timestamp>.xlsx
```

(unless `--inplace` or `--out` is specified)

---

## 🖥️ GUI Usage

For coworkers who don’t want the command line:

```bash
python run_gui_fill_updated.py
```

### GUI Features

* Pick **Template file**, **Period (YYYYMM)**, **Announce sheet**, and **Task**.
* Optional: override 2-3 / 4-3 sources, rates workbook, related-party master, etc.
* Choose **Overwrite template in-place** or set a custom **Output path**.
* One-click **Run** button.
* Live log panel shows what the script is doing.
* **Open Output Folder** button jumps straight to the results.
* Settings are remembered automatically in a `.state.json` file next to the GUI script.

---

## 📋 Typical Workflow

1. Download/update source files (`export_關係人交易-收入.xlsx`, `export_關係人交易-應收帳款.xlsx`, rates, MRS/RPTIS workbooks).
2. Launch GUI (`python run_gui_fill_updated.py`).
3. Select the **Template workbook** for the current month.
4. Enter the **Period (YYYYMM)**.
5. Choose **Task = all** (recommended).
6. Hit **Run**.
7. Check the **output folder**: `ytm_forms/data/output/`.

---

## ❗ Notes

* If external files are missing, the script will warn but still run (formulas may show `#REF!` in Excel).
* Excel external links (`RPTIS10`, `MRS0034`, `MRS0014`) must remain in their default locations unless overridden.
* If coworkers need to move files, update the overrides in the GUI.

---

## 👥 For Developers

* All logic lives in `ytm_forms/scripts/fill_updated.py`.
* GUI wrapper (`run_gui_fill_updated.py`) is just a thin Tkinter frontend → CLI.
* State is stored in `run_gui_fill_updated.state.json` (safe to delete anytime).

