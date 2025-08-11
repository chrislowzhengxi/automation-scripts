#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
# adjust import if your path differs
from fill_2_2 import run_mrs0014, run_rptis10

def browse_template():
    path = filedialog.askopenfilename(
        title="Choose template workbook",
        filetypes=[("Excel files", "*.xlsx")]
    )
    if path:
        template_var.set(path)

def run_task():
    task = task_var.get()
    period = period_var.get().strip()
    template = template_var.get().strip()

    if not period or len(period) != 6 or not period.isdigit():
        messagebox.showerror("Missing/invalid period", "Enter period as YYYYMM (e.g., 202504).")
        return
    if not template:
        messagebox.showerror("Missing template", "Please choose the template workbook (.xlsx).")
        return

    try:
        tpl = Path(template)
        # run in-place on the same template
        if task == "MRS0014":
            run_mrs0014(tpl, period, None, tpl)
        elif task == "RPTIS10":
            run_rptis10(tpl, period, None, tpl)
        else:  # Both
            run_mrs0014(tpl, period, None, tpl)
            run_rptis10(tpl, period, None, tpl)
        messagebox.showinfo("Done", f"{task} completed successfully.")
    except Exception as e:
        messagebox.showerror("Error", str(e))

root = tk.Tk()
root.title("Fill 2-2 (Click & Run)")
root.geometry("480x200")

task_var = tk.StringVar(value="MRS0014")
period_var = tk.StringVar()
template_var = tk.StringVar()

tk.Label(root, text="Task:").grid(row=0, column=0, sticky="e", padx=8, pady=8)
tk.OptionMenu(root, task_var, "MRS0014", "RPTIS10", "Both").grid(row=0, column=1, sticky="w", padx=8)

tk.Label(root, text="Period (YYYYMM):").grid(row=1, column=0, sticky="e", padx=8, pady=8)
tk.Entry(root, textvariable=period_var, width=20).grid(row=1, column=1, sticky="w", padx=8)

tk.Label(root, text="Template file:").grid(row=2, column=0, sticky="e", padx=8, pady=8)
tk.Entry(root, textvariable=template_var, width=36).grid(row=2, column=1, sticky="w", padx=8)
tk.Button(root, text="Browse...", command=browse_template).grid(row=2, column=2, padx=8)

tk.Button(root, text="Run", width=12, command=run_task).grid(row=3, column=1, pady=16)

root.mainloop()
