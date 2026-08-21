"""
Card Attendance Recorder  (runs on each TA's laptop, synced live via LAN)
============================================================================
Works with the OMNIKEY 5427 G2 in keyboard-emulation (HID) mode: the reader
"types" an 18-character hex string + Enter into whatever window has focus.
Only the last 10 characters (the fixed part) are kept as the card's UID.

TWO LAPTOPS, ONE HALL
----------------------
Two TAs each run this app on their own laptop with their own scanner, both
connected to the SAME mobile hotspot (a phone's hotspot, or a small travel
router) -- a private link the two of you control, independent of the
building's WiFi. Each laptop keeps a local SQLite database
(attendance_local.db) and they sync with each other directly over that
hotspot every ~12 seconds (or on demand via "Sync Now"), so both laptops
converge on the same picture of who's in/out even though two independent
scanners are recording it.

One laptop is the "Controller" (starts/ends sessions; this generates the
Session ID both laptops end up sharing) and the other is the "Helper"
(scans normally, learns the active session from the sync). Either laptop
can flip the "This laptop controls sessions" checkbox -- e.g. if the
Controller's laptop has a problem, the Helper ticks the box and takes over
for the rest of the day.

Requirements:
    pip install openpyxl

Usage:
    python attendance_recorder.py

Files used/created next to this script:
    students_database.db   -- read-only university roster (build with
                               build_student_database.py from the .xlsx
                               the university sends)
    attendance_local.db    -- this laptop's local database (roster/sessions/scans)
    device_config.json     -- this laptop's saved TA name / role / port
    attendance_logger.xlsx -- exported snapshot (auto-refreshed after each
                              sync and after End Session) for the later,
                              cross-hall merge + attendance_analyzer.py step
"""

import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import json
import os
import sys
from datetime import datetime

import attendance_db as db
import sync_service as sync


def _app_dir():
    """Folder the .exe (or .py, when not frozen) actually lives in -- NOT
    whatever the current working directory happens to be at launch, which
    can differ depending on how the app was started (double-click vs. a
    shortcut with a different 'Start in' path vs. a terminal)."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


APP_DIR = _app_dir()
DEVICE_CONFIG_FILE = os.path.join(APP_DIR, "device_config.json")
DB_FILE = os.path.join(APP_DIR, db.DB_FILE)
ATTENDANCE_EXPORT_FILE = os.path.join(APP_DIR, db.EXPORT_FILE)
STUDENT_DB_FILE = os.path.join(APP_DIR, "students_database.db")
DEBOUNCE_SECONDS = 3
UID_LENGTH = 10
TYPICAL_SESSION_RANGE = (45, 90)  # minutes; soft sanity check only
SYNC_INTERVAL_SECONDS = 12

# TODO: replace with the official list of halls once provided.
HALL_LIST = ["A.020", "A.128", "A.228", "A.328", "A.428"]

FACULTY_LIST = [
    "Biotechnology",
    "Business Administration",
    "Business Informatics",
    "Engineering",
    "Informatics and Computer Science",
    "Pharmaceutical Engineering",
]


def extract_uid(raw_scan):
    raw_scan = raw_scan.strip().upper()
    if len(raw_scan) >= UID_LENGTH:
        return raw_scan[-UID_LENGTH:]
    return raw_scan


def load_student_database(path=STUDENT_DB_FILE):
    """Read the university-provided roster from the read-only SQLite file
    built by build_student_database.py (table: students(id, name, faculty)).
    Opened with mode=ro so even a bug in this app couldn't write to it.
    Returns {student_id: {"name":.., "faculty":..}}, or {} if the file is
    missing (caller handles that -- manual enrollment still works)."""
    db_map = {}
    if not os.path.exists(path):
        return db_map
    import sqlite3

    uri = f"file:{os.path.abspath(path)}?mode=ro"
    try:
        conn = sqlite3.connect(uri, uri=True)
        try:
            for sid, name, faculty in conn.execute("SELECT id, name, faculty FROM students"):
                db_map[str(sid).strip()] = {"name": name or "", "faculty": faculty or ""}
        finally:
            conn.close()
    except sqlite3.OperationalError:
        return {}
    return db_map


def load_device_config():
    if os.path.exists(DEVICE_CONFIG_FILE):
        with open(DEVICE_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


def save_device_config(cfg):
    with open(DEVICE_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def ask_first_run_setup(root):
    """Blocking first-run dialog: TA name + controller role. Returns dict."""
    dialog = tk.Toplevel(root)
    dialog.title("First-time setup")
    dialog.transient(root)
    dialog.grab_set()
    result = {}

    ttk.Label(dialog, text="Your name (TA):").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 2))
    name_var = tk.StringVar()
    ttk.Entry(dialog, textvariable=name_var, width=32).grid(row=1, column=0, padx=10, pady=(0, 10))

    controller_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(
        dialog, text="This laptop controls sessions (Start/End Session)", variable=controller_var
    ).grid(row=2, column=0, sticky="w", padx=10, pady=(0, 10))
    ttk.Label(
        dialog,
        text="Only ONE of the two laptops should have this checked.\nThe other TA's laptop will follow along automatically.",
        justify="left",
        foreground="gray30",
    ).grid(row=3, column=0, sticky="w", padx=10, pady=(0, 10))

    def on_ok():
        if not name_var.get().strip():
            messagebox.showwarning("Missing name", "Enter your name.", parent=dialog)
            return
        result["ta_name"] = name_var.get().strip()
        result["is_controller"] = controller_var.get()
        dialog.destroy()

    ttk.Button(dialog, text="Start", command=on_ok).grid(row=4, column=0, pady=(0, 10))
    dialog.wait_window()
    return result


def ask_end_session_params(parent, title, initial_instructor="", initial_hall=None):
    """Modal dialog for (instructor name, duration in minutes, hall number)."""
    dialog = tk.Toplevel(parent)
    dialog.title(title)
    dialog.transient(parent)
    dialog.grab_set()
    result = {}

    ttk.Label(dialog, text="Instructor / TA name:").grid(row=0, column=0, sticky="w", padx=8, pady=(8, 2))
    instructor_var = tk.StringVar(value=initial_instructor)
    instructor_entry = ttk.Entry(dialog, textvariable=instructor_var, width=32)
    instructor_entry.grid(row=1, column=0, padx=8, pady=(0, 8))

    ttk.Label(dialog, text="Session duration (minutes):").grid(row=2, column=0, sticky="w", padx=8, pady=(0, 2))
    duration_var = tk.StringVar()
    duration_entry = ttk.Entry(dialog, textvariable=duration_var, width=32)
    duration_entry.grid(row=3, column=0, padx=8, pady=(0, 8))

    ttk.Label(dialog, text="Hall number:").grid(row=4, column=0, sticky="w", padx=8, pady=(0, 2))
    hall_var = tk.StringVar(value=initial_hall or (HALL_LIST[0] if HALL_LIST else ""))
    hall_combo = ttk.Combobox(dialog, textvariable=hall_var, values=HALL_LIST, state="readonly", width=30)
    hall_combo.grid(row=5, column=0, padx=8, pady=(0, 8))

    def on_ok():
        name = instructor_var.get().strip()
        dur_raw = duration_var.get().strip()
        hall = hall_var.get().strip()
        if not name:
            messagebox.showwarning("Missing name", "Enter the instructor/TA name.", parent=dialog)
            return
        if not hall:
            messagebox.showwarning("Missing hall", "Select the hall number.", parent=dialog)
            return
        try:
            duration = float(dur_raw)
            if duration <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning(
                "Invalid duration", "Enter the session duration in minutes as a positive number.", parent=dialog
            )
            return
        lo, hi = TYPICAL_SESSION_RANGE
        if not (lo <= duration <= hi):
            if not messagebox.askyesno(
                "Unusual duration",
                f"{duration:g} minutes is outside the typical {lo}-{hi} minute range. Continue anyway?",
                parent=dialog,
            ):
                return
        result["instructor"] = name
        result["duration"] = duration
        result["hall"] = hall
        dialog.destroy()

    btn_row = ttk.Frame(dialog)
    btn_row.grid(row=6, column=0, pady=(0, 8))
    ttk.Button(btn_row, text="OK", command=on_ok).pack(side="left", padx=4)
    ttk.Button(btn_row, text="Cancel", command=dialog.destroy).pack(side="left", padx=4)

    instructor_entry.focus_set()
    dialog.wait_window()
    if "duration" not in result:
        return None
    return result["instructor"], result["duration"], result["hall"]


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        db.init_db(DB_FILE)
        self.student_db = load_student_database()

        cfg = load_device_config()
        if cfg is None:
            cfg = ask_first_run_setup(self)
            if not cfg:
                messagebox.showinfo("Setup cancelled", "Setup was cancelled -- the application will now close.")
                self.destroy()
                return
            save_device_config(cfg)
        self.ta_name = cfg["ta_name"]
        self.is_controller = cfg["is_controller"]

        self.device_info = {"ta_name": self.ta_name, "is_controller": self.is_controller}
        self.sync_server = sync.start_server(DB_FILE, self.device_info, port=sync.DEFAULT_PORT)
        self.sync_manager = sync.SyncManager(DB_FILE)

        self.title(f"Card Attendance Recorder -- {self.ta_name}")
        self.geometry("820x640")

        self.last_scan = {}
        self.last_instructor = self.ta_name
        self.last_hall = None
        self.current_enroll_uid = None
        self._warned_both_controller = False

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.attendance_tab = ttk.Frame(self.notebook)
        self.enroll_tab = ttk.Frame(self.notebook)
        self.sync_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.attendance_tab, text="Attendance")
        self.notebook.add(self.enroll_tab, text="Enroll Cards")
        self.notebook.add(self.sync_tab, text="Setup & Sync")

        self._build_attendance_tab()
        self._build_enroll_tab()
        self._build_sync_tab()
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        self._refresh_session_ui()
        self.after(1000, self._poll_background_state)

        if not self.student_db:
            messagebox.showinfo(
                "Student database not found",
                f"'{STUDENT_DB_FILE}' wasn't found next to this script.\n"
                "If you have students_database.xlsx, run build_student_database.py "
                "once to convert it, then restart this app.\n\n"
                "Enroll Cards will still work with manual entry only until then.",
            )

    # ---------------- Attendance tab ----------------
    def _build_attendance_tab(self):
        frame = self.attendance_tab

        session_row = ttk.Frame(frame)
        session_row.pack(fill="x", padx=8, pady=8)
        self.start_btn = ttk.Button(session_row, text="Start Session", command=self._start_session)
        self.start_btn.pack(side="left")
        self.end_btn = ttk.Button(session_row, text="End Session", command=self._end_session)
        self.end_btn.pack(side="left", padx=(6, 0))
        self.session_status_label = ttk.Label(session_row, text="No active session", font=("Segoe UI", 10, "bold"))
        self.session_status_label.pack(side="left", padx=12)

        ttk.Label(frame, text="Scan a card (keep this tab open and this box focused):").pack(
            anchor="w", padx=8, pady=(4, 2)
        )
        self.att_entry = ttk.Entry(frame, font=("Consolas", 14))
        self.att_entry.pack(fill="x", padx=8)
        self.att_entry.bind("<Return>", self._on_attendance_scan)

        self.att_log = tk.Listbox(frame, font=("Consolas", 11))
        self.att_log.pack(fill="both", expand=True, padx=8, pady=8)

        self.att_status = ttk.Label(frame, text=f"Local database: {DB_FILE}")
        self.att_status.pack(anchor="w", padx=8, pady=(0, 8))

    def _on_attendance_scan(self, event):
        raw = self.att_entry.get()
        self.att_entry.delete(0, tk.END)
        uid = extract_uid(raw)
        if not uid:
            return

        roster_entry = db.get_roster_entry(uid, db_path=DB_FILE)
        if roster_entry is None:
            messagebox.showwarning(
                "Unregistered card",
                "This card is not registered in the system.\n"
                "Please add the student in the 'Enroll Cards' tab first.",
            )
            return

        open_session = db.get_open_session(db_path=DB_FILE)
        if open_session is None:
            messagebox.showwarning("No active session", "No session is currently running.")
            return

        self._register_attendance(uid, roster_entry, open_session)

    def _register_attendance(self, uid, roster_entry, open_session):
        now_key = datetime.now()
        last = self.last_scan.get(uid)
        if last and (now_key - last).total_seconds() < DEBOUNCE_SECONDS:
            return
        self.last_scan[uid] = now_key

        student_id = roster_entry.get("student_id", "")
        name = roster_entry.get("name", "UNKNOWN CARD")
        faculty = roster_entry.get("faculty", "")
        session_id = open_session["session_id"]
        ts = db.now_iso()
        scan_id = f"{self.ta_name}|{uid}|{ts}"

        db.insert_scan(scan_id, uid, student_id, name, faculty, session_id, self.ta_name, db_path=DB_FILE, ts=ts)

        line = f"{datetime.now().strftime('%H:%M:%S')}   {uid}   {student_id:<10}   {name:<20}   [{session_id}]"
        self.att_log.insert(0, line)

    # ---------------- Session controls ----------------
    def _start_session(self):
        if not self.is_controller:
            messagebox.showinfo("Helper laptop", "Only the Controller laptop starts sessions.")
            return
        if db.get_open_session(db_path=DB_FILE):
            messagebox.showinfo("Session already active", "A session is already running.")
            return

        today_nodash = datetime.now().strftime("%Y%m%d")
        n = db.highest_session_number(today_nodash, db_path=DB_FILE) + 1
        if n > 5:
            if not messagebox.askyesno(
                "More than 5 sessions today",
                f"This would be session {n} today (usual max is 5). Continue anyway?",
            ):
                return

        session_id = f"{today_nodash}_{n}"
        db.start_session(session_id, db_path=DB_FILE)
        self._refresh_session_ui()
        self.sync_manager.sync_now()  # push the new session to the partner right away

    def _end_session(self):
        if not self.is_controller:
            messagebox.showinfo("Helper laptop", "Only the Controller laptop ends sessions.")
            return
        open_session = db.get_open_session(db_path=DB_FILE)
        if not open_session:
            messagebox.showinfo("No active session", "There is no session currently running.")
            return

        session_id = open_session["session_id"]
        result = ask_end_session_params(self, f"End Session {session_id}", self.last_instructor, self.last_hall)
        if result is None:
            return
        instructor, duration, hall = result
        self.last_instructor = instructor
        self.last_hall = hall

        db.close_session(session_id, hall, instructor, duration, db_path=DB_FILE)
        self._refresh_session_ui()
        db.export_to_xlsx(db_path=DB_FILE, path=ATTENDANCE_EXPORT_FILE)
        self.sync_manager.sync_now()  # push the closure to the partner right away

        n_scans = sum(1 for s in db.get_all_scans(DB_FILE) if s["session_id"] == session_id)
        messagebox.showinfo(
            "Session closed",
            f"Session {session_id} closed.\n"
            f"Hall: {hall}   TA: {instructor}   Duration: {duration:g} min\n"
            f"{n_scans} scan row(s) recorded on this laptop (before syncing with the partner).\n\n"
            "Once every hall's export is merged, run attendance_analyzer.py to generate the report.",
        )

    def _refresh_session_ui(self):
        open_session = db.get_open_session(db_path=DB_FILE)
        if open_session:
            self.session_status_label.config(text=f"Active session: {open_session['session_id']}")
            self.att_entry.config(state="normal")
        else:
            self.session_status_label.config(text="No active session")
            self.att_entry.config(state="disabled")

        if self.is_controller:
            self.start_btn.config(state="disabled" if open_session else "normal")
            self.end_btn.config(state="normal" if open_session else "disabled")
        else:
            self.start_btn.config(state="disabled")
            self.end_btn.config(state="disabled")

    # ---------------- Enroll tab ----------------
    def _build_enroll_tab(self):
        frame = self.enroll_tab
        ttk.Label(
            frame,
            text="1) Scan the card   2) Type the Student ID and click Search\n"
            "3) Found -> Name/Faculty auto-fill (still editable)   Not found -> enter them manually\n"
            "If a session is running, enrolling also logs this moment as that student's scan.",
            justify="left",
        ).pack(anchor="w", padx=8, pady=(8, 4))

        uid_row = ttk.Frame(frame)
        uid_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(uid_row, text="Scan card:", width=14).pack(side="left")
        self.enroll_uid_var = tk.StringVar()
        self.enroll_uid_entry = ttk.Entry(uid_row, textvariable=self.enroll_uid_var, font=("Consolas", 12))
        self.enroll_uid_entry.pack(side="left", fill="x", expand=True)
        self.enroll_uid_entry.bind("<Return>", self._on_enroll_uid_scan)

        id_row = ttk.Frame(frame)
        id_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(id_row, text="Student ID:", width=14).pack(side="left")
        self.enroll_id_var = tk.StringVar()
        self.enroll_id_entry = ttk.Entry(id_row, textvariable=self.enroll_id_var, font=("Consolas", 12))
        self.enroll_id_entry.pack(side="left", fill="x", expand=True)
        self.enroll_id_entry.bind("<Return>", lambda e: self._on_search_student())
        ttk.Button(id_row, text="Search", command=self._on_search_student).pack(side="left", padx=(6, 0))

        self.enroll_search_status = ttk.Label(frame, text="")
        self.enroll_search_status.pack(anchor="w", padx=8)

        name_row = ttk.Frame(frame)
        name_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(name_row, text="Student name:", width=14).pack(side="left")
        self.enroll_name_var = tk.StringVar()
        self.enroll_name_entry = ttk.Entry(name_row, textvariable=self.enroll_name_var, font=("Consolas", 12))
        self.enroll_name_entry.pack(side="left", fill="x", expand=True)

        faculty_row = ttk.Frame(frame)
        faculty_row.pack(fill="x", padx=8, pady=4)
        ttk.Label(faculty_row, text="Faculty:", width=14).pack(side="left")
        self.enroll_faculty_var = tk.StringVar()
        self.enroll_faculty_combo = ttk.Combobox(
            faculty_row, textvariable=self.enroll_faculty_var, values=FACULTY_LIST, state="readonly", width=38
        )
        self.enroll_faculty_combo.pack(side="left", fill="x", expand=True)

        ttk.Button(frame, text="Enroll / Update Student", command=self._on_enroll_submit).pack(
            padx=8, pady=8, anchor="w"
        )

        self.enroll_list = tk.Listbox(frame, font=("Consolas", 11))
        self.enroll_list.pack(fill="both", expand=True, padx=8, pady=8)
        self._refresh_enroll_list()

        self.enroll_uid_entry.focus_set()

    def _on_enroll_uid_scan(self, event):
        raw = self.enroll_uid_var.get()
        uid = extract_uid(raw)
        if not uid:
            return
        self.current_enroll_uid = uid
        self.enroll_uid_var.set(uid)
        self.enroll_id_entry.focus_set()

    def _on_search_student(self):
        student_id = self.enroll_id_var.get().strip()
        if not student_id:
            messagebox.showwarning("Missing Student ID", "Type the Student ID first.")
            return

        match = self.student_db.get(student_id)
        if match:
            self.enroll_name_var.set(match["name"])
            self.enroll_faculty_var.set(match["faculty"])
            self.enroll_search_status.config(text="Found in the university database.", foreground="green")
        else:
            self.enroll_name_var.set("")
            self.enroll_faculty_var.set("")
            self.enroll_search_status.config(
                text="Not found in the database -- enter Name and Faculty manually.", foreground="red"
            )
        self.enroll_name_entry.focus_set()

    def _on_enroll_submit(self):
        uid = self.current_enroll_uid
        student_id = self.enroll_id_var.get().strip()
        name = self.enroll_name_var.get().strip()
        faculty = self.enroll_faculty_var.get().strip()

        if not uid:
            messagebox.showwarning("Missing card", "Scan a card in the 'Scan card' field first.")
            return
        if not student_id or not name or not faculty:
            messagebox.showwarning("Missing info", "Student ID, Name, and Faculty are all required.")
            return

        db.upsert_roster(uid, student_id, name, faculty, db_path=DB_FILE)
        self._refresh_enroll_list()

        open_session = db.get_open_session(db_path=DB_FILE)
        if open_session:
            self._register_attendance(uid, {"student_id": student_id, "name": name, "faculty": faculty}, open_session)
            note = f"Logged as attendance for session {open_session['session_id']}."
        else:
            note = "No active session right now, so attendance was not recorded."
        messagebox.showinfo("Student enrolled", f"{name} (ID {student_id}, {faculty}) enrolled.\n{note}")

        self.current_enroll_uid = None
        self.enroll_uid_var.set("")
        self.enroll_id_var.set("")
        self.enroll_name_var.set("")
        self.enroll_faculty_var.set("")
        self.enroll_search_status.config(text="")
        self.enroll_uid_entry.focus_set()

    def _refresh_enroll_list(self):
        self.enroll_list.delete(0, tk.END)
        for uid, info in db.get_roster(db_path=DB_FILE).items():
            self.enroll_list.insert(
                tk.END, f"{uid}   ID: {info['student_id']:<10}   {info['name']:<25}   {info.get('faculty', '')}"
            )

    # ---------------- Setup & Sync tab ----------------
    def _build_sync_tab(self):
        frame = self.sync_tab

        info_row = ttk.Frame(frame)
        info_row.pack(fill="x", padx=8, pady=8)
        my_ip = sync.get_local_ip()
        ttk.Label(
            info_row,
            text=f"This laptop -- TA: {self.ta_name}   |   My address: {my_ip}:{sync.DEFAULT_PORT}",
            font=("Segoe UI", 10, "bold"),
        ).pack(anchor="w")

        self.controller_var = tk.BooleanVar(value=self.is_controller)
        ttk.Checkbutton(
            frame,
            text="This laptop controls sessions (Start/End Session)",
            variable=self.controller_var,
            command=self._on_controller_toggle,
        ).pack(anchor="w", padx=8, pady=(4, 8))
        ttk.Label(
            frame,
            text="Only one laptop should have this checked. If the Controller's laptop has a\n"
            "problem, the other TA can check this box here to take over for the rest of the day.",
            justify="left",
            foreground="gray30",
        ).pack(anchor="w", padx=8, pady=(0, 12))

        ttk.Separator(frame).pack(fill="x", padx=8, pady=4)

        connect_row = ttk.Frame(frame)
        connect_row.pack(fill="x", padx=8, pady=8)
        ttk.Label(connect_row, text="Partner's address (shown on their screen):", width=32).pack(side="left")
        self.peer_ip_var = tk.StringVar()
        ttk.Entry(connect_row, textvariable=self.peer_ip_var, width=18).pack(side="left")
        ttk.Label(connect_row, text=f":{sync.DEFAULT_PORT}").pack(side="left")

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(btn_row, text="Connect & Start Syncing", command=self._on_connect).pack(side="left")
        ttk.Button(btn_row, text="Sync Now", command=self._on_sync_now).pack(side="left", padx=(6, 0))
        ttk.Button(btn_row, text="Disconnect", command=self._on_disconnect).pack(side="left", padx=(6, 0))

        self.sync_status_label = ttk.Label(frame, text="Not connected", foreground="gray30")
        self.sync_status_label.pack(anchor="w", padx=8, pady=(4, 12))

        ttk.Separator(frame).pack(fill="x", padx=8, pady=4)

        ttk.Button(frame, text="Export attendance_logger.xlsx now", command=self._on_manual_export).pack(
            anchor="w", padx=8, pady=8
        )
        ttk.Label(
            frame,
            text="(Also happens automatically after every successful sync and after End Session,\n"
            "so the export file is always close to current if you need it mid-class.)",
            justify="left",
            foreground="gray30",
        ).pack(anchor="w", padx=8)

    def _on_controller_toggle(self):
        self.is_controller = self.controller_var.get()
        self.device_info["is_controller"] = self.is_controller
        cfg = {"ta_name": self.ta_name, "is_controller": self.is_controller}
        save_device_config(cfg)
        self._refresh_session_ui()
        self._warned_both_controller = False

    def _on_connect(self):
        peer_ip = self.peer_ip_var.get().strip()
        if not peer_ip:
            messagebox.showwarning("Missing address", "Enter the partner laptop's address first.")
            return
        peer_url = f"http://{peer_ip}:{sync.DEFAULT_PORT}"
        try:
            reply = sync.ping(peer_url, timeout=4)
        except Exception as e:
            messagebox.showerror(
                "Couldn't reach partner",
                f"Could not reach {peer_ip}.\n\n"
                "Make sure both laptops are connected to the same mobile hotspot, "
                "the partner's app is running, and the address is typed correctly.\n\n"
                f"Details: {e}",
            )
            return
        if not reply.get("ok"):
            messagebox.showerror("Couldn't reach partner", "Partner responded unexpectedly. Try again.")
            return

        self.sync_manager.start(peer_url, interval_seconds=SYNC_INTERVAL_SECONDS)
        self.sync_status_label.config(text=f"Connected to {peer_ip} -- syncing every {SYNC_INTERVAL_SECONDS}s", foreground="green")

    def _on_sync_now(self):
        if not self.sync_manager.peer_base_url:
            messagebox.showinfo("Not connected", "Click 'Connect & Start Syncing' first.")
            return
        self.sync_manager.sync_now()
        self._apply_sync_status()

    def _on_disconnect(self):
        self.sync_manager.stop()
        self.sync_status_label.config(text="Not connected", foreground="gray30")

    def _on_manual_export(self):
        db.export_to_xlsx(db_path=DB_FILE, path=ATTENDANCE_EXPORT_FILE)
        messagebox.showinfo("Exported", f"Saved {ATTENDANCE_EXPORT_FILE}")

    # ---------------- Background state polling (main thread only) --------
    def _poll_background_state(self):
        self._apply_sync_status()
        self._refresh_session_ui()
        self._refresh_enroll_list()
        self.after(1000, self._poll_background_state)

    def _apply_sync_status(self):
        status = self.sync_manager.last_status
        color = "green" if status.lower().startswith("[") and "fail" not in status.lower() else "red"
        if self.sync_manager.peer_base_url is None:
            color = "gray30"
        self.sync_status_label.config(text=status, foreground=color)

        remote_info = self.sync_manager.last_remote_device_info
        if remote_info and remote_info.get("is_controller") and self.is_controller and not self._warned_both_controller:
            self._warned_both_controller = True
            messagebox.showwarning(
                "Both laptops set as Controller",
                "Your partner's laptop is also set to control sessions.\n"
                "Please uncheck it on one of the two laptops to avoid mismatched Session IDs.",
            )

        if self.sync_manager.peer_base_url:
            db.export_to_xlsx(db_path=DB_FILE, path=ATTENDANCE_EXPORT_FILE)

    # ---------------- Focus management ----------------
    def _on_tab_changed(self, event=None):
        try:
            current_tab = self.nametowidget(self.notebook.select())
            if current_tab == self.attendance_tab and str(self.att_entry.cget("state")) == "normal":
                self.att_entry.focus_set()
            elif current_tab == self.enroll_tab:
                self.enroll_uid_entry.focus_set()
        except Exception:
            pass


if __name__ == "__main__":
    app = App()
    if app.winfo_exists():
        app.mainloop()