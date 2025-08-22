import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from pathlib import Path
import subprocess, sys, os

# Paths relative to repo root
PROJECT_ROOT = Path(__file__).resolve().parents[1]
SCRIPTS_DIR = PROJECT_ROOT / "scripts"
OUTPUT_DIR = Path(__file__).resolve().parent.parent / "data" / "output"


def open_output_folder():
    if not OUTPUT_DIR.exists():
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    os.startfile(OUTPUT_DIR) 

def run_fill_updated(template, task, period, src43, src23, rates, relparty):
    cmd = [sys.executable, str(SCRIPTS_DIR / "fill_updated.py"),
           "--template", str(template),
           "--task", task,
           "--period", period]
    if src43: cmd += ["--src-43", str(src43)]
    if src23: cmd += ["--src-23", str(src23)]
    if rates: cmd += ["--rates-path", str(rates)]
    if relparty: cmd += ["--relparty-path", str(relparty)]
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True)
        return out
    except subprocess.CalledProcessError as e:
        return e.output

def run_fix_rates(period):
    cmd = [sys.executable, str(SCRIPTS_DIR / "fix_rates_ntd_to_usd.py"),
           "--period", period]
    try:
        out = subprocess.check_output(cmd, stderr=subprocess.STDOUT, text=True)
        return out
    except subprocess.CalledProcessError as e:
        return e.output

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("YTM Forms Automation")
        self.geometry("600x400")

        # Inputs
        frm = ttk.Frame(self); frm.pack(padx=10, pady=10, fill="x")

        ttk.Label(frm, text="Template File:").grid(row=0, column=0, sticky="w")
        self.template_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.template_var, width=50).grid(row=0, column=1)
        ttk.Button(frm, text="Browse", command=self.browse_template).grid(row=0, column=2)

        ttk.Label(frm, text="Task:").grid(row=1, column=0, sticky="w")
        self.task_var = tk.StringVar(value="both")
        ttk.Combobox(frm, textvariable=self.task_var,
                     values=["copy_4_3", "copy_2_3", "both"]).grid(row=1, column=1, sticky="w")

        ttk.Label(frm, text="Period (YYYYMM):").grid(row=2, column=0, sticky="w")
        self.period_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.period_var).grid(row=2, column=1, sticky="w")

        # optional overrides
        ttk.Label(frm, text="Rates workbook (optional):").grid(row=3, column=0, sticky="w")
        self.rates_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.rates_var, width=50).grid(row=3, column=1)
        ttk.Button(frm, text="Browse", command=lambda: self.browse_file(self.rates_var)).grid(row=3, column=2)

        ttk.Label(frm, text="Relparty workbook (optional):").grid(row=4, column=0, sticky="w")
        self.relparty_var = tk.StringVar()
        ttk.Entry(frm, textvariable=self.relparty_var, width=50).grid(row=4, column=1)
        ttk.Button(frm, text="Browse", command=lambda: self.browse_file(self.relparty_var)).grid(row=4, column=2)

        # Actions
        ttk.Button(frm, text="Run Fill Updated", command=self.do_fill).grid(row=5, column=1, pady=5)
        ttk.Button(frm, text="Fix NTD→USD", command=self.do_fix).grid(row=5, column=2, pady=5)

        # Output log
        self.log = tk.Text(self, wrap="word", height=10)
        self.log.pack(fill="both", expand=True, padx=10, pady=10)

        # Open Output folder 
        btn_open = ttk.Button(self, text="Open Output Folder", command=open_output_folder)
        btn_open.pack(pady=5)

    def browse_template(self):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if f: self.template_var.set(f)

    def browse_file(self, var):
        f = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls;*.xlsx")])
        if f: var.set(f)

    def do_fill(self):
        if not self.template_var.get() or not self.period_var.get():
            messagebox.showerror("Error", "Template and Period are required")
            return
        out = run_fill_updated(
            template=Path(self.template_var.get()),
            task=self.task_var.get(),
            period=self.period_var.get(),
            src43=None, src23=None,
            rates=self.rates_var.get() or None,
            relparty=self.relparty_var.get() or None
        )
        self.log.insert("end", out + "\n")
        self.log.see("end")

    def do_fix(self):
        if not self.period_var.get():
            messagebox.showerror("Error", "Period is required")
            return
        out = run_fix_rates(self.period_var.get())
        self.log.insert("end", out + "\n")
        self.log.see("end")

if __name__ == "__main__":
    App().mainloop()
