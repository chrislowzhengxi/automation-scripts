#!/usr/bin/env python3
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
from datetime import datetime
import shutil

# Import your existing functions
from fill import run_mrs0014, run_rptis10, run_export_paste

APP_TITLE = "Fill 2-2 & 4-3 (All-in-One)"
DEFAULT_DEST_SHEET_43 = "4-3.應收關係人科餘"
FIXED_DEST_START_ROW_43 = 10

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

def do_backup_if_checked(template_path: Path) -> Path:
    """If 'Backup before overwrite' is checked, create a timestamped copy and return its path.
       Otherwise return the original template path (in-place)."""
    if not backup_var.get():
        return template_path
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"BACKUP_{ts}_{template_path.name}"
    backup_path = template_path.with_name(backup_name)
    shutil.copy(template_path, backup_path)
    return backup_path

def on_mode_change(*_):
    mode = mode_var.get()
    is_43 = (mode == "4-3")
    # Toggle visibility for 2-2 vs 4-3 fields
    two2_frame.grid() if mode == "2-2" else two2_frame.grid_remove()
    four3_frame.grid() if is_43 else four3_frame.grid_remove()

def run_clicked():
    mode = mode_var.get()
    period = period_var.get().strip()
    template = template_var.get().strip()

    # Common validations
    if not (period.isdigit() and len(period) == 6):
        messagebox.showerror("Missing/invalid period", "Enter period as YYYYMM (e.g., 202504).")
        return
    if not template:
        messagebox.showerror("Missing template", "Please choose the template workbook (.xlsx).")
        return

    tpl = Path(template)
    if not tpl.exists():
        messagebox.showerror("Template not found", f"File not found:\n{tpl}")
        return

    try:
        if mode == "2-2":
            subtask = two2_task_var.get()  # "MRS0014", "RPTIS10", "Both"
            out_target = do_backup_if_checked(tpl)

            if subtask == "MRS0014":
                run_mrs0014(out_target, period, None, out_target)
            elif subtask == "RPTIS10":
                run_rptis10(out_target, period, None, out_target, src_sheet=None)
            else:  # Both
                run_mrs0014(out_target, period, None, out_target)
                run_rptis10(out_target, period, None, out_target, src_sheet=None)

            messagebox.showinfo("Done", f"2-2 ({subtask}) completed.\nWritten to:\n{out_target}")

        else:  # 4-3
            export_path = export_var.get().strip()
            dest_sheet = dest_sheet_var.get().strip() or DEFAULT_DEST_SHEET_43

            if not export_path:
                messagebox.showerror("Missing export", "Please choose the export.xlsx file.")
                return

            exp = Path(export_path)
            if not exp.exists():
                messagebox.showerror("Export not found", f"File not found:\n{exp}")
                return

            out_target = do_backup_if_checked(tpl)
            run_export_paste(out_target, exp, out_target,
                             dest_sheet=dest_sheet,
                             dst_start_row=FIXED_DEST_START_ROW_43)

            messagebox.showinfo("Done", f"4-3 export completed.\nWritten to:\n{out_target}")

    except Exception as e:
        messagebox.showerror("Error", str(e))

# ---------------- UI ----------------
root = tk.Tk()
root.title(APP_TITLE)
root.geometry("640x360")

# Shared controls
mode_var = tk.StringVar(value="2-2")           # "2-2" or "4-3"
period_var = tk.StringVar()
template_var = tk.StringVar()
backup_var = tk.BooleanVar(value=True)         # default: backup ON

# Top row: Mode selector
top_row = tk.Frame(root)
top_row.grid(row=0, column=0, sticky="we", padx=10, pady=(10, 2))
tk.Label(top_row, text="Mode:").grid(row=0, column=0, sticky="e", padx=(0,6))
tk.OptionMenu(top_row, mode_var, "2-2", "4-3", command=lambda _: on_mode_change()).grid(row=0, column=1, sticky="w")

# Shared fields
shared = tk.LabelFrame(root, text="Shared")
shared.grid(row=1, column=0, sticky="we", padx=10, pady=6)
shared.grid_columnconfigure(1, weight=1)

tk.Label(shared, text="Period (YYYYMM):").grid(row=0, column=0, sticky="e", padx=8, pady=6)
tk.Entry(shared, textvariable=period_var, width=20).grid(row=0, column=1, sticky="w", padx=6, pady=6)

tk.Label(shared, text="Template file:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
tk.Entry(shared, textvariable=template_var, width=44).grid(row=1, column=1, sticky="we", padx=6, pady=6)
tk.Button(shared, text="Browse...", command=browse_template).grid(row=1, column=2, padx=6, pady=6, sticky="w")

tk.Checkbutton(shared, text="Backup before overwrite", variable=backup_var).grid(row=2, column=1, sticky="w", padx=6, pady=(0,6))

# 2-2 section
two2_frame = tk.LabelFrame(root, text="2-2 Options")
two2_frame.grid(row=2, column=0, sticky="we", padx=10, pady=6)
two2_frame.grid_columnconfigure(1, weight=1)

two2_task_var = tk.StringVar(value="Both")
tk.Label(two2_frame, text="Task:").grid(row=0, column=0, sticky="e", padx=8, pady=8)
tk.OptionMenu(two2_frame, two2_task_var, "MRS0014", "RPTIS10", "Both").grid(row=0, column=1, sticky="w", padx=6, pady=8)

# 4-3 section
four3_frame = tk.LabelFrame(root, text="4-3 Options")
four3_frame.grid(row=3, column=0, sticky="we", padx=10, pady=6)
four3_frame.grid_columnconfigure(1, weight=1)

export_var = tk.StringVar()
dest_sheet_var = tk.StringVar(value=DEFAULT_DEST_SHEET_43)

tk.Label(four3_frame, text="Export file:").grid(row=0, column=0, sticky="e", padx=8, pady=6)
tk.Entry(four3_frame, textvariable=export_var, width=44).grid(row=0, column=1, sticky="we", padx=6, pady=6)
tk.Button(four3_frame, text="Browse...", command=browse_export).grid(row=0, column=2, padx=6, pady=6, sticky="w")

tk.Label(four3_frame, text="Dest. sheet:").grid(row=1, column=0, sticky="e", padx=8, pady=6)
tk.Entry(four3_frame, textvariable=dest_sheet_var, width=44).grid(row=1, column=1, sticky="we", padx=6, pady=6)

# Bottom: Run button
run_row = tk.Frame(root)
run_row.grid(row=4, column=0, sticky="e", padx=10, pady=10)
tk.Button(run_row, text="Run", width=16, command=run_clicked).grid(row=0, column=0, sticky="e")

# Initialize visibility
on_mode_change()

root.mainloop()
