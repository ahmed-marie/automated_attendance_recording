"""
Attendance Analyzer  (Phase II - run ONCE, after all halls' logs are merged)
=============================================================================
Each hall runs attendance_recorder.py locally and produces its own
attendance_logger.xlsx. At day's end, every hall's file is merged into one
(e.g. uploaded into a shared Google Sheet, then downloaded back as a single
attendance_logger.xlsx). This script is then run ONCE, by the lecturer or a
senior TA, on that merged file to produce the day's attendance report.

Merged file columns (per scan row):
    Time | UID | Student ID | Name | Session ID | Hall Number | TA Name | Duration (min)

Status rules, evaluated per (student, session_id) -- note the SAME session_id
can legitimately appear under multiple Hall Numbers (each hall runs its own
"Session 1", "Session 2", ...), so hall consistency is checked explicitly:

    - The student's scans for that session_id span MORE THAN ONE Hall Number
        -> "Needs Review - Hall Conflict"
    - Otherwise, an odd number of scans for that student+session
        -> "Needs Review - Scan Missing"   (same rule as before, just renamed)
    - Otherwise, pair scans as (in, out), (in, out)... and sum the time
        -> "Attended"  if total minutes >= session_duration x ATTENDANCE_RATIO
        -> "Absent"    otherwise

A session's duration is taken from the Duration (min) column itself (the
most common value entered for that session_id, across all its rows/halls)
-- there's no separate metadata file to keep in sync, so re-running this
script after fixing a cell in Excel is all "re-evaluating" requires.

Requirements:
    pip install openpyxl

Usage:
    python attendance_analyzer.py
    python attendance_analyzer.py --input attendance_logger.xlsx --date 2026-08-19
"""

import argparse
import json
import os
from collections import Counter
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter

ATTENDANCE_FILE_DEFAULT = "attendance_logger.xlsx"
ROSTER_FILE = "roster.json"
REPORT_PREFIX = "attendance_report_"

# A student must be inside the hall for at least this fraction of the
# session's actual (entered) duration.
ATTENDANCE_RATIO = 0.75


def load_roster(path=ROSTER_FILE):
    if os.path.exists(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def session_number(session_id):
    """'20260819_2' -> 2"""
    return int(session_id.rsplit("_", 1)[1])


def read_day_scans(recorder_path, date_str):
    """Return {session_id: {uid: {"student_id","name","times":[...],
    "halls": set(...), "durations": [...]}}} for every scan on date_str,
    across every hall present in the merged file."""
    sessions = {}
    if not os.path.exists(recorder_path):
        return sessions

    wb = openpyxl.load_workbook(recorder_path, data_only=True)
    if date_str not in wb.sheetnames:
        return sessions
    ws = wb[date_str]
    sheet_date = datetime.strptime(date_str, "%Y-%m-%d").date()

    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None:
            continue
        padded = (row + (None,) * 8)[:8]
        time_val, uid, student_id, name, session_id, hall, ta_name, duration = padded
        if uid is None or session_id is None:
            continue
        uid = str(uid).strip()

        if isinstance(time_val, datetime):
            ts = time_val
        elif hasattr(time_val, "hour"):  # datetime.time
            ts = datetime.combine(sheet_date, time_val)
        else:
            ts = datetime.combine(sheet_date, datetime.strptime(str(time_val), "%H:%M:%S").time())

        sess = sessions.setdefault(session_id, {})
        entry = sess.setdefault(
            uid, {"student_id": student_id or "", "name": name or "", "times": [], "halls": set(), "durations": []}
        )
        if student_id:
            entry["student_id"] = student_id
        if name and name != "UNKNOWN CARD":
            entry["name"] = name
        entry["times"].append(ts)
        if hall:
            entry["halls"].add(hall)
        if duration not in (None, ""):
            try:
                entry["durations"].append(float(duration))
            except (TypeError, ValueError):
                pass

    return sessions


def session_duration(sess_students):
    """Most common Duration value entered across this session's rows.
    Returns None if no Duration was ever recorded for this session."""
    all_durations = [d for info in sess_students.values() for d in info["durations"]]
    if not all_durations:
        return None
    return Counter(all_durations).most_common(1)[0][0]


def evaluate_student(info, threshold_minutes):
    if len(info["halls"]) > 1:
        return "Needs Review - Hall Conflict"

    times = sorted(info["times"])
    n = len(times)
    if n % 2 != 0:
        return "Needs Review - Scan Missing"

    total_seconds = 0.0
    for i in range(0, n, 2):
        total_seconds += (times[i + 1] - times[i]).total_seconds()
    total_minutes = total_seconds / 60.0
    return "Attended" if total_minutes >= threshold_minutes else "Absent"


def generate_report(recorder_path, date_str, roster, report_path=None):
    sessions = read_day_scans(recorder_path, date_str)
    ordered_ids = sorted(sessions.keys(), key=session_number)

    durations = {sid: session_duration(sessions[sid]) for sid in ordered_ids}

    warnings = []
    for sid in ordered_ids:
        distinct = sorted({d for info in sessions[sid].values() for d in info["durations"]})
        if len(distinct) > 1:
            warnings.append(
                f"Session {sid}: inconsistent durations entered across halls {distinct}, using {durations[sid]:g} min."
            )

    all_uids = set(roster.keys())
    for sid in ordered_ids:
        all_uids |= set(sessions[sid].keys())

    def label(uid):
        if uid in roster:
            return roster[uid].get("student_id", uid), roster[uid].get("name", "UNKNOWN CARD")
        for sid in ordered_ids:
            info = sessions[sid].get(uid)
            if info and info["name"]:
                return info["student_id"], info["name"]
        return uid, "UNKNOWN CARD"

    rows = []
    for uid in sorted(all_uids):
        student_id, name = label(uid)
        row = {"uid": uid, "student_id": student_id, "name": name, "statuses": {}}
        for sid in ordered_ids:
            dur = durations[sid]
            info = sessions[sid].get(uid, {"times": [], "halls": set(), "durations": []})
            if dur is None:
                row["statuses"][sid] = "Needs Review - Scan Missing" if info["times"] else "Absent"
            else:
                row["statuses"][sid] = evaluate_student(info, dur * ATTENDANCE_RATIO)
        rows.append(row)

    rows.sort(key=lambda r: (r["name"] == "UNKNOWN CARD", r["student_id"]))

    report_path = report_path or f"{REPORT_PREFIX}{date_str}.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Attendance Report"

    ws.append([f"Attendance Report - {date_str}"])
    for idx, sid in enumerate(ordered_ids, start=1):
        dur = durations[sid]
        halls_used = sorted({h for info in sessions[sid].values() for h in info["halls"]})
        halls_note = f" (Halls: {', '.join(halls_used)})" if halls_used else ""
        dur_note = f"{dur:g} minutes" if dur is not None else "unknown (no Duration recorded)"
        ws.append([f"Session {idx} Duration: {dur_note}{halls_note}"])
    ws.append([])

    header = ["Card UID", "Student ID", "Name"] + ordered_ids
    ws.append(header)
    for row in rows:
        line = [row["uid"], row["student_id"], row["name"]] + [row["statuses"][sid] for sid in ordered_ids]
        ws.append(line)

    widths = {"A": 16, "B": 14, "C": 26}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for i in range(len(ordered_ids)):
        ws.column_dimensions[get_column_letter(4 + i)].width = 22

    wb.save(report_path)
    return report_path, rows, ordered_ids, warnings


def main():
    p = argparse.ArgumentParser(description="Generate the daily attendance report from the merged log.")
    p.add_argument("--input", default=ATTENDANCE_FILE_DEFAULT, help="Path to the (merged) attendance_logger.xlsx")
    p.add_argument("--date", default=datetime.now().strftime("%Y-%m-%d"), help="Report date, YYYY-MM-DD")
    p.add_argument("--output", default=None, help="Output report path (default: attendance_report_<date>.xlsx)")
    args = p.parse_args()

    if not os.path.exists(args.input):
        raise SystemExit(f"Input file not found: {args.input}")

    roster = load_roster()
    report_path, rows, ordered_ids, warnings = generate_report(args.input, args.date, roster, args.output)

    counts = {}
    for row in rows:
        for status in row["statuses"].values():
            counts[status] = counts.get(status, 0) + 1

    print(f"Report saved to: {report_path}")
    print("Totals:", ", ".join(f"{k}: {v}" for k, v in counts.items()) if counts else "no data found")
    for w in warnings:
        print(f"WARNING: {w}")


if __name__ == "__main__":
    main()
