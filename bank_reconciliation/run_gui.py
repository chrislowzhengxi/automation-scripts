import os
import sys
import threading
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox

APP_TITLE = "Bank Reconciliation – Runner"

def validate_date(s: str) -> bool:
    if len(s) != 8 or not s.isdigit():
        return False
    try:
        datetime.strptime(s, "%Y%m%d")
        return True
    except ValueError:
        return False

class BankRunnerGUI:
    def __init__(self, master: tk.Tk):
        self.master = master
        master.title(APP_TITLE)
        master.geometry("820x600")

        top = tk.Frame(master)
        top.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(top, text="Posting date (YYYYMMDD):").pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.today().strftime("%Y%m%d"))
        self.date_entry = tk.Entry(top, textvariable=self.date_var, width=12)
        self.date_entry.pack(side=tk.LEFT, padx=(6, 10))

        tk.Button(top, text="Today", command=self.set_today).pack(side=tk.LEFT)
        tk.Button(top, text="Add files…", command=self.add_files).pack(side=tk.LEFT, padx=(20, 6))
        tk.Button(top, text="Remove selected", command=self.remove_selected).pack(side=tk.LEFT)
        tk.Button(top, text="Clear", command=self.clear_files).pack(side=tk.LEFT, padx=(6, 0))

        mid = tk.Frame(master)
        mid.pack(fill=tk.BOTH, expand=True, padx=10)
        tk.Label(mid, text="Files to process:").pack(anchor="w")
        self.file_list = tk.Listbox(mid, selectmode=tk.EXTENDED)
        self.file_list.pack(fill=tk.BOTH, expand=True, pady=(4, 10))

        runrow = tk.Frame(master)
        runrow.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(runrow, text="Run", command=self.run_clicked).pack(side=tk.LEFT)
        tk.Button(runrow, text="Open output folder", command=self.open_output_folder).pack(side=tk.LEFT, padx=10)

        logframe = tk.Frame(master)
        logframe.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        tk.Label(logframe, text="Log:").pack(anchor="w")
        self.log = tk.Text(logframe, height=16, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll = tk.Scrollbar(logframe, command=self.log.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.config(yscrollcommand=scroll.set)

        self.status_var = tk.StringVar(value="Ready")
        status = tk.Label(master, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W)
        status.pack(fill=tk.X, side=tk.BOTTOM)

    def set_today(self):
        self.date_var.set(datetime.today().strftime("%Y%m%d"))

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select bank statement files",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        for p in paths:
            if p and p not in self.file_list.get(0, tk.END):
                self.file_list.insert(tk.END, p)

    def remove_selected(self):
        for i in reversed(self.file_list.curselection()):
            self.file_list.delete(i)

    def clear_files(self):
        self.file_list.delete(0, tk.END)

    def open_output_folder(self):
        downloads = os.path.join(os.path.expanduser("~"), "Downloads")
        banks = os.path.join(downloads, "Banks")
        target = banks if os.path.isdir(banks) else downloads
        try:
            os.startfile(target)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{target}\n\n{e}")

    def append_log(self, text: str):
        self.log.configure(state=tk.NORMAL)
        self.log.insert(tk.END, text)
        self.log.see(tk.END)
        self.log.configure(state=tk.DISABLED)

    def set_status(self, text: str):
        self.status_var.set(text)

    def run_clicked(self):
        files = list(self.file_list.get(0, tk.END))
        if not files:
            messagebox.showwarning("No files", "Please add at least one Excel file.")
            return
        ymd = self.date_var.get().strip()
        if not validate_date(ymd):
            messagebox.showerror("Invalid date", "Please enter a valid date in YYYYMMDD format.")
            return

        script_dir = os.path.dirname(os.path.abspath(__file__))
        bank_py = os.path.join(script_dir, "bank.py")
        if not os.path.isfile(bank_py):
            messagebox.showerror("Missing bank.py", f"bank.py not found next to this app:\n{bank_py}")
            return

        self.set_status("Running…")
        self.append_log("\n=== Run started ===\n")
        self.master.after(100, lambda: self._toggle_run_buttons(False))
        t = threading.Thread(target=self._run_all, args=(bank_py, files, ymd), daemon=True)
        t.start()

    def _toggle_run_buttons(self, enable: bool):
        for btn in self.master.winfo_children():
            if isinstance(btn, tk.Frame):
                for sub in btn.winfo_children():
                    if isinstance(sub, tk.Button):
                        sub.configure(state=(tk.NORMAL if enable else tk.DISABLED))

    def _run_all(self, bank_py: str, files: list[str], ymd: str):
        for f in files:
            self._run_single(bank_py, f, ymd)
        self.master.after(0, lambda: self._toggle_run_buttons(True))
        self.master.after(0, lambda: self.set_status("Done"))
        self.master.after(0, lambda: self.append_log("=== Run finished ===\n"))

    def _run_single(self, bank_py: str, filepath: str, ymd: str):
        self.master.after(0, lambda: self.append_log(f"\n-- Processing: {os.path.basename(filepath)} --\n"))
        cmd = [sys.executable, bank_py, "-f", filepath, "-d", ymd]
        try:
            proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                universal_newlines=True,
                bufsize=1,
            )
            for line in proc.stdout:
                self.master.after(0, lambda s=line: self.append_log(s))
            proc.wait()
            rc = proc.returncode
            if rc != 0:
                self.master.after(0, lambda: self.append_log(f"[ERROR] Process exited with code {rc}\n"))
            else:
                self.master.after(0, lambda: self.append_log("[OK] Completed.\n"))
        except Exception as e:
            self.master.after(0, lambda: self.append_log(f"[EXCEPTION] {e}\n"))

def main():
    root = tk.Tk()
    app = BankRunnerGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()