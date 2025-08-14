#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path

# Import your function directly
from fill import run_export_paste

def browse_template():
    path = filedialog.askopenfilename(
        title="Choose template workbook",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        template_var.set(path)

def browse_export():
    path = filedialog.askopenfilename(
        title="Choose export.xlsx file",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        export_var.set(path)

def run_task():
    period = period_var.get().strip()
    template = template_var.get().strip()
    export_path = export_var.get().strip()
    dest_sheet = dest_sheet_var.get().strip()
    dest_row = 10  # fixed per your requirement

    if not (period.isdigit() and len(period) == 6):
        messagebox.showerror("Missing/invalid period", "Enter period as YYYYMM (e.g., 202504).")
        return
    if not template:
        messagebox.showerror("Missing template", "Please choose the template workbook (.xlsx).")
        return
    if not export_path:
        messagebox.showerror("Missing export", "Please choose the export.xlsx file.")
        return
    if not dest_sheet:
        messagebox.showerror("Missing sheet", "Please specify the destination sheet name.")
        return

    try:
        tpl = Path(template)
        exp = Path(export_path)
        run_export_paste(tpl, exp, tpl,
                         dest_sheet=dest_sheet,
                         dst_start_row=dest_row)
        messagebox.showinfo("Done", f"4-3 export completed successfully.\nWritten to: {tpl}")
    except Exception as e:
        messagebox.showerror("Error", str(e))

# ---------------- UI ----------------
root = tk.Tk()
root.title("Fill 4-3 Export")
root.geometry("560x240")

period_var = tk.StringVar()
template_var = tk.StringVar()
export_var = tk.StringVar()
dest_sheet_var = tk.StringVar(value="4-3.應收關係人科餘")

# Row 0: Sheet label
tk.Label(root, text="Sheet:").grid(row=0, column=0, sticky="e", padx=8, pady=10)
tk.Label(root, text="4-3 Export").grid(row=0, column=1, sticky="w", padx=(0,6), pady=10)

# Row 1: Period
tk.Label(root, text="Period (YYYYMM):").grid(row=1, column=0, sticky="e", padx=8, pady=8)
tk.Entry(root, textvariable=period_var, width=20).grid(row=1, column=1, columnspan=2, sticky="w", padx=6, pady=8)

# Row 2: Template file
tk.Label(root, text="Template file:").grid(row=2, column=0, sticky="e", padx=8, pady=8)
tk.Entry(root, textvariable=template_var, width=36).grid(row=2, column=1, sticky="w", padx=6, pady=8)
tk.Button(root, text="Browse...", command=browse_template).grid(row=2, column=2, padx=6, pady=8, sticky="w")

# Row 3: Export file
tk.Label(root, text="Export file:").grid(row=3, column=0, sticky="e", padx=8, pady=8)
tk.Entry(root, textvariable=export_var, width=36).grid(row=3, column=1, sticky="w", padx=6, pady=8)
tk.Button(root, text="Browse...", command=browse_export).grid(row=3, column=2, padx=6, pady=8, sticky="w")

# Row 4: Destination sheet
tk.Label(root, text="Dest. sheet:").grid(row=4, column=0, sticky="e", padx=8, pady=8)
tk.Entry(root, textvariable=dest_sheet_var, width=36).grid(row=4, column=1, columnspan=2, sticky="w", padx=6, pady=8)

# Row 5: Run button
tk.Button(root, text="Run", width=14, command=run_task).grid(row=5, column=1, pady=16, sticky="e")

root.mainloop()
