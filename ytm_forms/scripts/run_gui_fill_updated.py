#!/usr/bin/env python3
"""
Tkinter GUI wrapper for
    python -m ytm_forms.scripts.fill_updated

Features
- Required fields: template, period (YYYYMM), task, announce sheet
- Optional overrides: src-43, src-23, rates, relparty, rptis10, mrs0034, mrs0014, out path, inplace toggle
- Live log window (captures stdout/stderr)
- Open output folder shortcut (defaults to PROJECT_ROOT/ytm_forms/data/output)
- Remembers last inputs in a small JSON next to this script

Place this file anywhere inside your repo. It auto-detects PROJECT_ROOT by
searching upward for a folder that contains "ytm_forms".
"""
from __future__ import annotations
import json
import os
import queue
import subprocess
import sys
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

APP_TITLE = "YTM Forms – GUI for fill_updated"
STATE_FILE = Path(__file__).with_suffix(".state.json")

# ---------- Repo / paths detection ----------

def find_project_root(start: Path) -> Path:
    cur = start.resolve()
    for parent in [cur, *cur.parents]:
        if (parent / "ytm_forms").is_dir():
            return parent
    # Fallback: assume script is two levels under root
    return start.resolve().parents[1]

PROJECT_ROOT = find_project_root(Path(__file__))
OUTPUT_DIR = PROJECT_ROOT / "ytm_forms" / "data" / "output"
TEMPLATE_DIR = PROJECT_ROOT / "ytm_forms" / "data" / "template"
REL_DEFAULT_DIR = TEMPLATE_DIR / "關係人"

# ---------- State helpers ----------

def load_state() -> dict:
    try:
        with open(STATE_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}

def save_state(d: dict) -> None:
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(d, f, ensure_ascii=False, indent=2)
    except Exception:
        pass

# ---------- GUI ----------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("860x640")
        self.minsize(820, 600)
        self.state = load_state()

        self._build_vars()
        self._build_ui()
        self.proc: subprocess.Popen | None = None
        self.log_queue: queue.Queue[str] = queue.Queue()
        self._pump_logs()

    # --- Vars ---
    def _build_vars(self):
        self.template_var = tk.StringVar(value=self.state.get("template", ""))
        self.period_var = tk.StringVar(value=self.state.get("period", "202504"))
        self.announce_var = tk.StringVar(value=self.state.get("announce", "1-1.公告(元)"))
        self.task_var = tk.StringVar(value=self.state.get("task", "all"))
        self.inplace_var = tk.BooleanVar(value=self.state.get("inplace", False))
        self.out_var = tk.StringVar(value=self.state.get("out", ""))
        # Advanced
        self.src43_var = tk.StringVar(value=self.state.get("src43", ""))
        self.src23_var = tk.StringVar(value=self.state.get("src23", ""))
        self.rates_var = tk.StringVar(value=self.state.get("rates", ""))
        self.relparty_var = tk.StringVar(value=self.state.get("relparty", ""))
        self.rptis10_var = tk.StringVar(value=self.state.get("rptis10", ""))
        self.mrs0034_var = tk.StringVar(value=self.state.get("mrs0034", ""))
        self.mrs0014_var = tk.StringVar(value=self.state.get("mrs0014", ""))

    # --- UI ---
    def _build_ui(self):
        pad = {"padx": 8, "pady": 6}

        # Top: Required group
        req = ttk.LabelFrame(self, text="Required")
        req.pack(fill="x", **pad)

        # Template
        row = ttk.Frame(req)
        row.pack(fill="x", **pad)
        ttk.Label(row, text="Template .xlsx:").pack(side="left")
        ent = ttk.Entry(row, textvariable=self.template_var)
        ent.pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row, text="Browse", command=self._pick_template).pack(side="left")

        # Period + Announce sheet
        row = ttk.Frame(req)
        row.pack(fill="x", **pad)
        ttk.Label(row, text="Period (YYYYMM):").pack(side="left")
        ttk.Entry(row, width=10, textvariable=self.period_var).pack(side="left", padx=8)
        ttk.Label(row, text="Announce sheet:").pack(side="left", padx=(16, 0))
        ttk.Entry(row, width=28, textvariable=self.announce_var).pack(side="left", padx=8)

        # Task
        row = ttk.Frame(req)
        row.pack(fill="x", **pad)
        ttk.Label(row, text="Task:").pack(side="left")
        for val, lbl in [
            ("all", "All"),
            ("both", "Copy 4-3 + 2-3"),
            ("copy_4_3", "Copy 4-3 only"),
            ("copy_2_3", "Copy 2-3 only"),
            ("announce_structure", "Announce structure only"),
        ]:
            ttk.Radiobutton(row, text=lbl, value=val, variable=self.task_var).pack(side="left", padx=6)

        # Output / inplace
        row = ttk.Frame(req)
        row.pack(fill="x", **pad)
        ttk.Checkbutton(row, text="Overwrite template (in-place)", variable=self.inplace_var, command=self._toggle_out).pack(side="left")
        ttk.Label(row, text="Output path (optional):").pack(side="left", padx=(16, 0))
        out_ent = ttk.Entry(row, textvariable=self.out_var)
        out_ent.pack(side="left", fill="x", expand=True, padx=8)
        self.out_entry = out_ent
        ttk.Button(row, text="Save As…", command=self._pick_out).pack(side="left")
        self._toggle_out()

        # Advanced group
        adv = ttk.LabelFrame(self, text="Advanced (optional overrides)")
        adv.pack(fill="x", **pad)

        self._path_row(adv, "4-3 source (export_關係人交易-應收帳款.xlsx):", self.src43_var, self._pick_src43, REL_DEFAULT_DIR)
        self._path_row(adv, "2-3 source (export_關係人交易-收入.xlsx):", self.src23_var, self._pick_src23, REL_DEFAULT_DIR)
        self._path_row(adv, "Rates workbook (… Ending 及 Avg … .xls):", self.rates_var, self._pick_rates, REL_DEFAULT_DIR)
        self._path_row(adv, "Related-party master (關係企業(人).xls):", self.relparty_var, self._pick_relparty, REL_DEFAULT_DIR)
        self._path_row(adv, "RPTIS10 workbook:", self.rptis10_var, self._pick_rptis10, REL_DEFAULT_DIR)
        self._path_row(adv, "MRS0034 workbook:", self.mrs0034_var, self._pick_mrs0034, REL_DEFAULT_DIR)
        self._path_row(adv, "MRS0014 workbook:", self.mrs0014_var, self._pick_mrs0014, REL_DEFAULT_DIR)

        # Actions
        actions = ttk.Frame(self)
        actions.pack(fill="x", **pad)
        ttk.Button(actions, text="Run", command=self._run).pack(side="left")
        ttk.Button(actions, text="Open Output Folder", command=self._open_output_folder).pack(side="left", padx=8)
        ttk.Button(actions, text="Clear Log", command=self._clear_log).pack(side="left")

        ttk.Label(actions, text=f"Output default: {OUTPUT_DIR}").pack(side="right")

        # Log
        logf = ttk.LabelFrame(self, text="Log")
        logf.pack(fill="both", expand=True, **pad)
        self.log = tk.Text(logf, height=18, wrap="none")
        self.log.pack(fill="both", expand=True)
        self.log.configure(state="disabled")

    def _path_row(self, parent, label, var, cb, initial_dir: Path):
        row = ttk.Frame(parent)
        row.pack(fill="x", padx=8, pady=3)
        ttk.Label(row, text=label).pack(side="left")
        ent = ttk.Entry(row, textvariable=var)
        ent.pack(side="left", fill="x", expand=True, padx=8)
        def pick():
            cb(initial_dir)
        ttk.Button(row, text="Browse", command=pick).pack(side="left")

    # --- Browsers ---
    def _pick_template(self):
        path = filedialog.askopenfilename(title="Choose template workbook", initialdir=TEMPLATE_DIR, filetypes=[("Excel", "*.xlsx")])
        if path:
            self.template_var.set(path)

    def _pick_out(self):
        path = filedialog.asksaveasfilename(title="Choose output workbook path", defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")])
        if path:
            self.out_var.set(path)

    def _pick_src43(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose 4-3 source", initialdir=initial_dir, filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.src43_var.set(p)

    def _pick_src23(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose 2-3 source", initialdir=initial_dir, filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.src23_var.set(p)

    def _pick_rates(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose rates workbook", initialdir=initial_dir, filetypes=[("Excel", "*.xls;*.xlsx")])
        if p:
            self.rates_var.set(p)

    def _pick_relparty(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose related-party master", initialdir=initial_dir, filetypes=[("Excel", "*.xls;*.xlsx")])
        if p:
            self.relparty_var.set(p)

    def _pick_rptis10(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose RPTIS10 workbook", initialdir=initial_dir, filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.rptis10_var.set(p)

    def _pick_mrs0034(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose MRS0034 workbook", initialdir=initial_dir, filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.mrs0034_var.set(p)

    def _pick_mrs0014(self, initial_dir):
        p = filedialog.askopenfilename(title="Choose MRS0014 workbook", initialdir=initial_dir, filetypes=[("Excel", "*.xlsx;*.xls")])
        if p:
            self.mrs0014_var.set(p)

    # --- Actions ---
    def _toggle_out(self):
        if self.inplace_var.get():
            self.out_entry.configure(state="disabled")
        else:
            self.out_entry.configure(state="normal")

    def _open_output_folder(self):
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
        try:
            os.startfile(str(OUTPUT_DIR))  # Windows
        except Exception:
            messagebox.showinfo("Open", f"Output folder: {OUTPUT_DIR}")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", tk.END)
        self.log.configure(state="disabled")

    def _append_log(self, text: str):
        self.log.configure(state="normal")
        self.log.insert(tk.END, text)
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _validate(self) -> bool:
        tmpl = Path(self.template_var.get().strip())
        per = self.period_var.get().strip()
        if not tmpl.exists():
            messagebox.showerror("Missing", "Please choose a valid Template .xlsx file.")
            return False
        if not per.isdigit() or len(per) != 6:
            messagebox.showerror("Invalid period", "Period must be 6 digits: YYYYMM")
            return False
        return True

    def _collect_args(self) -> list[str]:
        args = [
            sys.executable, "-m", "ytm_forms.scripts.fill_updated",
            "--template", self.template_var.get().strip(),
            "--task", self.task_var.get(),
            "--period", self.period_var.get().strip(),
            "--announce-sheet", self.announce_var.get().strip() or "1-1.公告(元)",
        ]
        if self.inplace_var.get():
            args.append("--inplace")
        outp = self.out_var.get().strip()
        if outp and not self.inplace_var.get():
            args += ["--out", outp]
        # Optional overrides
        for flag, var in [
            ("--src-43", self.src43_var),
            ("--src-23", self.src23_var),
            ("--rates-path", self.rates_var),
            ("--relparty-path", self.relparty_var),
            ("--rptis10-path", self.rptis10_var),
            ("--mrs0034-path", self.mrs0034_var),
            ("--mrs0014-path", self.mrs0014_var),
        ]:
            v = var.get().strip()
            if v:
                args += [flag, v]
        return args

    def _run(self):
        if self.proc is not None:
            messagebox.showwarning("Busy", "A run is already in progress.")
            return
        if not self._validate():
            return

        # Save state
        save_state({
            "template": self.template_var.get(),
            "period": self.period_var.get(),
            "announce": self.announce_var.get(),
            "task": self.task_var.get(),
            "inplace": self.inplace_var.get(),
            "out": self.out_var.get(),
            "src43": self.src43_var.get(),
            "src23": self.src23_var.get(),
            "rates": self.rates_var.get(),
            "relparty": self.relparty_var.get(),
            "rptis10": self.rptis10_var.get(),
            "mrs0034": self.mrs0034_var.get(),
            "mrs0014": self.mrs0014_var.get(),
        })

        args = self._collect_args()
        self._append_log("\n=== Run started ===\n")
        self._append_log("[CMD] " + " ".join(self._quote_args(args)) + "\n")

        # Launch in PROJECT_ROOT to ensure module import works
        def target():
            try:
                self.proc = subprocess.Popen(
                    args,
                    cwd=str(PROJECT_ROOT),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    bufsize=1,
                )
                # Stream output
                assert self.proc.stdout is not None
                for line in self.proc.stdout:
                    self.log_queue.put(line)
                rc = self.proc.wait()
                self.log_queue.put(f"[EXIT] Code {rc}\n")
            except FileNotFoundError as e:
                self.log_queue.put(f"[ERROR] {e}\n")
            except Exception as e:
                self.log_queue.put(f"[ERROR] {type(e).__name__}: {e}\n")
            finally:
                self.proc = None

        threading.Thread(target=target, daemon=True).start()

    def _pump_logs(self):
        try:
            while True:
                line = self.log_queue.get_nowait()
                self._append_log(line)
        except queue.Empty:
            pass
        self.after(80, self._pump_logs)

    @staticmethod
    def _quote_args(args: list[str]) -> list[str]:
        out = []
        for a in args:
            if " " in a or "(" in a or ")" in a:
                out.append(f'"{a}"')
            else:
                out.append(a)
        return out


if __name__ == "__main__":
    try:
        OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    app = App()
    app.mainloop()
