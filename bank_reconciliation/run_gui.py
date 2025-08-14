import os
import sys
import threading
import subprocess
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

APP_TITLE = "Bank Reconciliation – Runner"

# If bank.py imports other deps, add them here so the GUI checks early.
REQUIRED_MODULES = [
    "rapidfuzz",   # fuzzy matching used by bank.py/fuzzy_matcher.py
    "pandas", "openpyxl", "xlrd",  # uncomment if you want GUI to verify these too
]

# ----- Utilities -----
def validate_date(s: str) -> bool:
    if len(s) != 8 or not s.isdigit():
        return False
    try:
        datetime.strptime(s, "%Y%m%d")
        return True
    except ValueError:
        return False

def check_dependencies():
    missing = []
    for mod in REQUIRED_MODULES:
        try:
            __import__(mod)
        except ImportError:
            missing.append(mod)
    return missing

class BankRunnerGUI:
    def __init__(self, master: tk.Tk):
        self.master = master
        master.title(APP_TITLE)
        master.geometry("840x640")

        # Top controls
        top = tk.Frame(master)
        top.pack(fill=tk.X, padx=10, pady=10)

        tk.Label(top, text="Posting date (YYYYMMDD):").pack(side=tk.LEFT)
        self.date_var = tk.StringVar(value=datetime.today().strftime("%Y%m%d"))
        tk.Entry(top, textvariable=self.date_var, width=12).pack(side=tk.LEFT, padx=(6, 10))

        tk.Button(top, text="Today", command=self.set_today).pack(side=tk.LEFT)
        tk.Button(top, text="Add files…", command=self.add_files).pack(side=tk.LEFT, padx=(20, 6))
        tk.Button(top, text="Remove selected", command=self.remove_selected).pack(side=tk.LEFT)
        tk.Button(top, text="Clear", command=self.clear_files).pack(side=tk.LEFT, padx=(6, 0))

        # # New checkbox
        self.new_run_var = tk.BooleanVar(value=False)
        tk.Checkbutton(
            top,
            text="Start new run (-2/-3…) for this batch",
            variable=self.new_run_var
        ).pack(side=tk.LEFT, padx=(20, 0))

        # File list
        mid = tk.Frame(master)
        mid.pack(fill=tk.BOTH, expand=True, padx=10)
        tk.Label(mid, text="Files to process:").pack(anchor="w")
        self.file_list = tk.Listbox(mid, selectmode=tk.EXTENDED)
        self.file_list.pack(fill=tk.BOTH, expand=True, pady=(4, 10))

        # Run row
        runrow = tk.Frame(master)
        runrow.pack(fill=tk.X, padx=10, pady=(0, 10))
        tk.Button(runrow, text="Run", command=self.run_clicked).pack(side=tk.LEFT)
        tk.Button(runrow, text="Open output folder", command=self.open_output_folder).pack(side=tk.LEFT, padx=10)

        # Log
        logframe = tk.Frame(master)
        logframe.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))
        tk.Label(logframe, text="Log:").pack(anchor="w")
        self.log = tk.Text(logframe, height=18, state=tk.DISABLED)
        self.log.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        scroll = tk.Scrollbar(logframe, command=self.log.yview)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.log.config(yscrollcommand=scroll.set)

        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        tk.Label(master, textvariable=self.status_var, bd=1, relief=tk.SUNKEN, anchor=tk.W).pack(fill=tk.X, side=tk.BOTTOM)


    # More helpers
    def _ask_on_main(self, fn):
        import threading
        done = threading.Event()
        box = {"val": None}
        def run():
            try:
                box["val"] = fn()
            finally:
                done.set()
        self.master.after(0, run)
        done.wait()
        return box["val"]

    def _ask_yes_no_sync(self, question: str) -> bool:
        return bool(self._ask_on_main(lambda:
            messagebox.askyesno("Confirm match", question, parent=self.master)
        ))

    def _ask_text_sync(self, question: str) -> str | None:
        return self._ask_on_main(lambda:
            simpledialog.askstring("Manual ID", question, parent=self.master)
        )


    # ----- UI helpers -----
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

    # ----- Run flow -----
    def run_clicked(self):
        files = list(self.file_list.get(0, tk.END))
        if not files:
            messagebox.showwarning("No files", "Please add at least one Excel file.")
            return

        ymd = self.date_var.get().strip()
        if not validate_date(ymd):
            messagebox.showerror("Invalid date", "Please enter a valid date in YYYYMMDD format.")
            return

        # Show which Python will be used (helps diagnose venv vs system Python)
        self.append_log(f"\nPython interpreter: {sys.executable}\n")
        self.append_log(f"sys.path[0]: {sys.path[0]}\n")

        # Verify deps are installed **in this interpreter**
        missing = check_dependencies()
        if missing:
            pip_cmd = f"{sys.executable} -m pip install {' '.join(missing)}"
            messagebox.showerror(
                "Missing dependencies",
                "The following modules are required but not installed in this Python:\n"
                + ", ".join(missing)
                + f"\n\nRun this command in Command Prompt:\n    {pip_cmd}"
            )
            self.append_log(f"[MISSING] {', '.join(missing)}\nTry: {pip_cmd}\n")
            return

        script_dir = os.path.dirname(os.path.abspath(__file__))
        bank_py = os.path.join(script_dir, "bank.py")
        if not os.path.isfile(bank_py):
            messagebox.showerror("Missing bank.py", f"bank.py not found next to this app:\n{bank_py}")
            return

        # self.set_status("Running…")
        # self.append_log("\n=== Run started ===\n")
        # self.master.after(100, lambda: self._toggle_run_buttons(False))
        # threading.Thread(target=self._run_all, args=(bank_py, files, ymd), daemon=True).start()
        self.set_status("Running…")
        self.append_log("\n=== Run started ===\n")

        # NEW: capture checkbox value and log the mode
        batch_new_run = bool(self.new_run_var.get())
        self.append_log(f"[MODE] {'NEW RUN' if batch_new_run else 'Append to latest'} for {ymd}\n")

        self.master.after(100, lambda: self._toggle_run_buttons(False))
        threading.Thread(
            target=self._run_all,
            args=(bank_py, files, ymd, batch_new_run),  # <- pass 4th arg
            daemon=True
        ).start()

    def _toggle_run_buttons(self, enable: bool):
        for child in self.master.winfo_children():
            if isinstance(child, tk.Frame):
                for sub in child.winfo_children():
                    if isinstance(sub, tk.Button):
                        sub.configure(state=(tk.NORMAL if enable else tk.DISABLED))

    # def _run_all(self, bank_py: str, files: list[str], ymd: str):
    #     for f in files:
    #         self._run_single(bank_py, f, ymd)
    #     self.master.after(0, lambda: self._toggle_run_buttons(True))
    #     self.master.after(0, lambda: self.set_status("Done"))
    #     self.master.after(0, lambda: self.append_log("=== Run finished ===\n"))
    def _run_all(self, bank_py: str, files: list[str], ymd: str, batch_new_run: bool):
        first = True
        for f in files:
            use_new_run = (first and batch_new_run)
            self._run_single(bank_py, f, ymd, new_run=use_new_run)
            first = False
        self.master.after(0, lambda: self._toggle_run_buttons(True))
        self.master.after(0, lambda: self.set_status("Done"))
        self.master.after(0, lambda: self.append_log("=== Run finished ===\n"))


    # def _run_single(self, bank_py: str, filepath: str, ymd: str):
    #     self.master.after(0, lambda: self.append_log(f"\n-- Processing: {os.path.basename(filepath)} --\n"))
    #     cmd = [sys.executable, bank_py, "-f", filepath, "-d", ymd]

    #     try:
    #         env = dict(os.environ)
    #         env["PYTHONIOENCODING"] = "utf-8"

    #         proc = subprocess.Popen(
    #             cmd,
    #             stdin=subprocess.PIPE,                 # <— allow sending answers
    #             stdout=subprocess.PIPE,
    #             stderr=subprocess.STDOUT,
    #             text=True,
    #             encoding="utf-8",
    #             errors="replace",
    #             bufsize=1,
    #             env=env,
    #         )
    def _run_single(self, bank_py: str, filepath: str, ymd: str, new_run: bool = False):
        self.master.after(0, lambda: self.append_log(f"\n-- Processing: {os.path.basename(filepath)} --\n"))
        cmd = [sys.executable, bank_py, "-f", filepath, "-d", ymd]
        if new_run:
            cmd.append("--new-run")
            self.master.after(0, lambda: self.append_log("[INFO] Starting NEW RUN for this batch (will write to -N file)\n"))

        # NEW: echo the exact command (great for debugging)
        self.master.after(0, lambda s=f"[CMD] {' '.join(cmd)}\n": self.append_log(s))

        try:
            env = dict(os.environ)
            env["PYTHONIOENCODING"] = "utf-8"

            proc = subprocess.Popen(
                cmd,
                stdin=subprocess.PIPE,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                bufsize=1,
                env=env,
            )
            for line in proc.stdout:
                # Handle interactive prompts
                if line.startswith("[[PROMPT:YN]]"):
                    q = line.split("]]", 1)[1].strip()
                    ans = self._ask_yes_no_sync(q)
                    to_send = "y" if ans else "n"
                    proc.stdin.write(to_send + "\n"); proc.stdin.flush()
                    self.master.after(0, lambda s=f"[UI] {q} → {to_send}\n": self.append_log(s))
                    continue

                if line.startswith("[[PROMPT:TEXT]]"):
                    q = line.split("]]", 1)[1].strip()
                    ans = self._ask_text_sync(q)
                    to_send = "" if ans in (None, "") else str(ans)
                    proc.stdin.write(to_send + "\n"); proc.stdin.flush()
                    shown = to_send if to_send else "(blank)"
                    self.master.after(0, lambda s=f"[UI] {q} → {shown}\n": self.append_log(s))
                    continue

                # Normal output
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