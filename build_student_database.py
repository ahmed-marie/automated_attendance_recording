"""
build_student_database.py
===========================
One-time (or "whenever the university sends an updated roster") conversion
tool: turns the university-provided students_database.xlsx into
students_database.db (SQLite). attendance_recorder.py then reads that .db
file READ-ONLY, so the ~400-student roster can't be accidentally edited by
a non-technical TA -- they only ever interact with the recorder app, never
this spreadsheet-turned-database directly.

Usage:
    python build_student_database.py
    python build_student_database.py --input students_database.xlsx --output students_database.db

The resulting .db is also marked read-only on disk as a second layer of
protection. Re-running this script (e.g. after the university sends an
updated roster) clears that flag, rebuilds the file from scratch, and
re-applies read-only automatically -- you don't need to unlock it by hand.

Requirements:
    pip install openpyxl
"""

import argparse
import os
import sqlite3
import stat

import openpyxl


def build(input_path, output_path):
    if not os.path.exists(input_path):
        raise SystemExit(f"Input file not found: {input_path}")

    if os.path.exists(output_path):
        os.chmod(output_path, stat.S_IWRITE | stat.S_IREAD)  # clear read-only so we can overwrite
        os.remove(output_path)

    wb = openpyxl.load_workbook(input_path, data_only=True)
    ws = wb.active

    conn = sqlite3.connect(output_path)
    conn.execute("CREATE TABLE students (id TEXT PRIMARY KEY, name TEXT NOT NULL, faculty TEXT NOT NULL)")

    imported, skipped = 0, 0
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        padded = (row + (None,) * 3)[:3]
        sid, name, faculty = padded
        sid = str(sid).strip()
        if not sid:
            continue
        name = str(name).strip() if name is not None else ""
        faculty = str(faculty).strip() if faculty is not None else ""
        if not name or not faculty:
            skipped += 1
            continue
        conn.execute("INSERT OR REPLACE INTO students (id, name, faculty) VALUES (?, ?, ?)", (sid, name, faculty))
        imported += 1

    conn.commit()
    conn.close()

    os.chmod(output_path, stat.S_IREAD)  # read-only on disk -- extra guard beyond the app opening it read-only

    print(f"Built {output_path}: {imported} students imported, {skipped} row(s) skipped (missing Name or Faculty).")
    if skipped:
        print("Skipped rows are missing a Name or Faculty value -- worth checking the source spreadsheet.")


def main():
    p = argparse.ArgumentParser(description="Convert students_database.xlsx into a read-only students_database.db")
    p.add_argument("--input", default="students_database.xlsx")
    p.add_argument("--output", default="students_database.db")
    args = p.parse_args()
    build(args.input, args.output)


if __name__ == "__main__":
    main()