"""
attendance_analyzer_gui.py
=============================
Desktop GUI for attendance_analyzer.py, for the lecturer/senior TA who
runs Phase II report generation and would rather not use a command line.

Browse to the merged attendance_logger.xlsx (it can live anywhere), pick
which date(s) to include, and generate the report. The report is always
saved NEXT TO THIS APPLICATION -- not next to the input file -- so every
generated report collects in one predictable place regardless of where
TAs' exported logs happen to be sitting.

Does NOT need students_database.db. That file is only used by
attendance_recorder.py's Enroll Cards search -- this analyzer's roster
info comes entirely from the "Roster" sheet embedded inside
attendance_logger.xlsx itself, written automatically by
attendance_recorder.py's export.

Requirements:
    pip install openpyxl

Usage:
    python attendance_analyzer_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import subprocess
import platform

import attendance_analyzer as analyzer


def _app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _app_dir()


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Report Generator")
        self.geometry("640x580")

        self.input_path_var = tk.StringVar()
        self.last_report_path = None

        self._build_ui()

    def _build_ui(self):
        pad = {"padx": 10, "pady": 6}

        ttk.Label(self, text="Attendance log file (attendance_logger.xlsx):").pack(anchor="w", **pad)
        path_row = ttk.Frame(self)
        path_row.pack(fill="x", padx=10)
        self.path_entry = ttk.Entry(path_row, textvariable=self.input_path_var, state="readonly")
        self.path_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(path_row, text="Browse...", command=self._on_browse).pack(side="left", padx=(6, 0))

        ttk.Label(self, text="Dates to include (all selected by default -- click to narrow down):").pack(
            anchor="w", **pad
        )
        list_frame = ttk.Frame(self)
        list_frame.pack(fill="both", expand=False, padx=10)
        self.date_listbox = tk.Listbox(list_frame, selectmode="extended", height=8, exportselection=False)
        self.date_listbox.pack(side="left", fill="both", expand=True)
        scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=self.date_listbox.yview)
        scrollbar.pack(side="left", fill="y")
        self.date_listbox.config(yscrollcommand=scrollbar.set)

        btn_row = ttk.Frame(self)
        btn_row.pack(fill="x", padx=10, pady=(6, 10))
        self.generate_btn = ttk.Button(btn_row, text="Generate Report", command=self._on_generate, state="disabled")
        self.generate_btn.pack(side="left")
        self.open_btn = ttk.Button(btn_row, text="Open Report", command=self._on_open_report, state="disabled")
        self.open_btn.pack(side="left", padx=(6, 0))

        ttk.Separator(self).pack(fill="x", padx=10, pady=4)

        ttk.Label(self, text="Status:").pack(anchor="w", padx=10)
        self.status_text = tk.Text(self, height=14, wrap="word", state="disabled")
        self.status_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        ttk.Label(
            self,
            text=f"Reports are always saved next to this application:\n{APP_DIR}",
            foreground="gray30",
            wraplength=600,
            justify="left",
        ).pack(anchor="w", padx=10, pady=(0, 10))

    def _log(self, text, clear=False):
        self.status_text.config(state="normal")
        if clear:
            self.status_text.delete("1.0", tk.END)
        self.status_text.insert(tk.END, text + "\n")
        self.status_text.config(state="disabled")
        self.status_text.see(tk.END)

    def _on_browse(self):
        path = filedialog.askopenfilename(
            title="Select attendance_logger.xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_path_var.set(path)
        self._load_dates(path)

    def _load_dates(self, path):
        self.date_listbox.delete(0, tk.END)
        self.generate_btn.config(state="disabled")
        try:
            dates = analyzer.list_available_dates(path)
        except Exception as e:
            messagebox.showerror("Couldn't read file", f"Couldn't read dates from this file:\n{e}")
            return

        if not dates:
            messagebox.showwarning(
                "No scan data found",
                "No rows with a recognizable Session ID (format YYYYMMDD_N) were found in this file.",
            )
            return

        for d in dates:
            self.date_listbox.insert(tk.END, d)
        self.date_listbox.select_set(0, tk.END)  # default: every date selected
        self.generate_btn.config(state="normal")
        self._log(f"Found {len(dates)} date(s): {', '.join(dates)}", clear=True)

    def _on_generate(self):
        input_path = self.input_path_var.get().strip()
        if not input_path or not os.path.exists(input_path):
            messagebox.showwarning("No file selected", "Choose an attendance_logger.xlsx file first.")
            return

        selected_indices = self.date_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("No dates selected", "Select at least one date to include.")
            return
        selected_dates = [self.date_listbox.get(i) for i in selected_indices]

        try:
            roster = analyzer.load_roster(input_path)
            report_path, per_date_rows, warnings, counts = analyzer.generate_full_report(
                input_path,
                roster,
                dates=selected_dates,
                report_path=self._default_output_path(selected_dates),
            )
        except PermissionError:
            messagebox.showerror(
                "File is open",
                "Couldn't save the report -- it looks like it's already open in Excel. Close it and try again.",
            )
            return
        except Exception as e:
            messagebox.showerror("Couldn't generate report", str(e))
            return

        self.last_report_path = report_path
        self.open_btn.config(state="normal")

        self._log(f"Report saved to: {report_path}", clear=True)
        self._log(f"Dates covered: {', '.join(per_date_rows.keys())}")
        totals = ", ".join(f"{k}: {v}" for k, v in counts.items()) if counts else "no data found"
        self._log(f"Totals: {totals}")
        for w in warnings:
            self._log(f"WARNING: {w}")

        messagebox.showinfo("Report generated", f"Saved to:\n{report_path}")

    def _default_output_path(self, selected_dates):
        dates = sorted(selected_dates)
        if len(dates) == 1:
            fname = f"{analyzer.REPORT_PREFIX}{dates[0]}.xlsx"
        else:
            fname = f"{analyzer.REPORT_PREFIX}{dates[0]}_to_{dates[-1]}.xlsx"
        return os.path.join(APP_DIR, fname)

    def _on_open_report(self):
        if not self.last_report_path or not os.path.exists(self.last_report_path):
            messagebox.showwarning("No report yet", "Generate a report first.")
            return
        try:
            if platform.system() == "Windows":
                os.startfile(self.last_report_path)
            elif platform.system() == "Darwin":
                subprocess.run(["open", self.last_report_path])
            else:
                subprocess.run(["xdg-open", self.last_report_path])
        except Exception as e:
            messagebox.showerror("Couldn't open file", str(e))


if __name__ == "__main__":
    App().mainloop()
