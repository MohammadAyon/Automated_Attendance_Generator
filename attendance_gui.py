"""
attendance_gui.py
------------------
Minimal GUI for the attendance automation: pick the raw punch-log CSV,
pick where to save the generated workbook, click Generate.

Requires: openpyxl  (pip install openpyxl)
Tkinter ships with Python on Windows/macOS. On Linux, if the import fails:
    sudo apt install python3-tk

Run with:
    python3 attendance_gui.py

This file must sit in the same folder as attendance_core.py.

To change the standard work hours, weekly off day, or this month's public
holidays, edit the constants at the top of attendance_core.py -- there's no
way to detect holidays automatically from the punch data, so the GUI does
not ask for them; they live in code so they're easy to find and edit.
"""

import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

import attendance_core as core


class AttendanceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Employee Attendance Generator")
        self.geometry("620x480")
        self.resizable(False, False)

        self.csv_path = tk.StringVar()
        self.output_path = tk.StringVar()

        tk.Label(self, text="Employee Attendance Generator",
                 font=("Arial", 16, "bold")).pack(pady=(16, 4))
        tk.Label(self, text="Turns the raw punch-log CSV into the formatted "
                             "workbook (Settings, Normalized Data, Summary, "
                             "one tab per employee).",
                 fg="#555", wraplength=560, justify="center").pack(pady=(0, 12))

        self._file_row("1. Raw punch-log CSV file:", self.csv_path,
                        self._browse_csv)
        self._file_row("2. Save the generated workbook as:", self.output_path,
                        self._browse_save)

        self.generate_btn = tk.Button(self, text="Generate Report",
                                       font=("Arial", 12, "bold"),
                                       bg="#2d6cdf", fg="white",
                                       command=self.on_generate)
        self.generate_btn.pack(pady=16)

        self.progress = ttk.Progressbar(self, mode="indeterminate", length=560)
        self.progress.pack(pady=(0, 10))

        tk.Label(self, text="Log:").pack(anchor="w", padx=12)
        self.log_box = scrolledtext.ScrolledText(self, height=13, width=74,
                                                   state="disabled",
                                                   font=("Consolas", 9))
        self.log_box.pack(padx=12, pady=(0, 12))

    def _file_row(self, label, var, browse_command):
        frame = tk.Frame(self)
        frame.pack(fill="x", padx=12, pady=6)
        tk.Label(frame, text=label).pack(anchor="w")
        row = tk.Frame(frame)
        row.pack(fill="x")
        tk.Entry(row, textvariable=var, width=58).pack(side="left", fill="x", expand=True)
        tk.Button(row, text="Browse...", command=browse_command).pack(side="left", padx=(6, 0))

    def _browse_csv(self):
        path = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if not path:
            return
        self.csv_path.set(path)
        if not self.output_path.get():
            self._suggest_output_name(path)

    def _suggest_output_name(self, csv_path):
        # Best-effort: peek at the CSV to name the file after the detected month.
        try:
            employees = core.parse_csv(csv_path)
            year, month = core.detect_month(employees)
            import calendar
            name = f"Attendance_Automated_{calendar.month_abbr[month]}{year}.xlsx"
        except Exception:
            name = "Attendance_Automated.xlsx"
        default_dir = str(Path(csv_path).parent)
        self.output_path.set(str(Path(default_dir) / name))

    def _browse_save(self):
        initial = Path(self.output_path.get()).name if self.output_path.get() else "Attendance_Automated.xlsx"
        path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                             filetypes=[("Excel files", "*.xlsx")],
                                             initialfile=initial)
        if path:
            self.output_path.set(path)

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", message + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def on_generate(self):
        csv_path = self.csv_path.get().strip()
        output_path = self.output_path.get().strip()

        if not csv_path or not Path(csv_path).exists():
            messagebox.showerror("Missing file", "Please choose a valid punch-log CSV file.")
            return
        if not output_path:
            messagebox.showerror("Missing output", "Please choose where to save the generated workbook.")
            return

        self.generate_btn.configure(state="disabled")
        self.progress.start(12)
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

        thread = threading.Thread(target=self._run_generation,
                                   args=(csv_path, output_path), daemon=True)
        thread.start()

    def _run_generation(self, csv_path, output_path):
        try:
            self.log(f"Reading {csv_path} ...")
            stats = core.build_workbook(csv_path, output_path, log=self.log)

            self.log("")
            self.log(f"Done. Report period: {stats['month']}")
            self.log(f"Employee tabs created: {stats['employees']}")
            if stats["zero_punch_employees"]:
                self.log(f"No punch data at all this month for: "
                          f"{', '.join(stats['zero_punch_employees'])} "
                          "(worth checking these are real employees, not device artifacts).")
            if stats["single_punch_days"]:
                self.log(f"Single-punch days (missing a check-in or check-out): "
                          f"{stats['single_punch_days']} -- marked 'Incomplete' in Remarks.")
            if stats["on_leave_days"]:
                self.log(f"Working days marked 'On Leave' (no punch record): {stats['on_leave_days']}")
            if stats["early_checkins"]:
                self.log(f"Check-ins before 4am on {len(stats['early_checkins'])} day(s) -- "
                          "could be a real night shift or a stray duplicate scan:")
                for emp, d, t in stats["early_checkins"][:10]:
                    self.log(f"    {emp}  {d}  {t}")

            self.log("")
            self.log(f"Saved to: {output_path}")
            self.after(0, lambda: messagebox.showinfo("Done", f"Workbook saved to:\n{output_path}"))

        except Exception as e:
            self.log("")
            self.log("ERROR: " + str(e))
            self.log(traceback.format_exc())
            self.after(0, lambda: messagebox.showerror("Generation failed", str(e)))
        finally:
            self.after(0, self._finish)

    def _finish(self):
        self.progress.stop()
        self.generate_btn.configure(state="normal")


if __name__ == "__main__":
    app = AttendanceApp()
    app.mainloop()
