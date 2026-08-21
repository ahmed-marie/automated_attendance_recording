# automated_attendance_recording

Two-phase card-based attendance system for lecture halls:

- `attendance_recorder.py` -- runs on each hall's PC during class, logging card
  scans from an OMNIKEY 5427 G2 (keyboard-emulation mode) to a local
  `attendance_logger.xlsx`.
- `attendance_analyzer.py` -- run once, after every hall's log file is merged
  into a single spreadsheet, to produce the day's attendance report.

## Setup

Requires Python 3.9+ with Tkinter (bundled with the standard installer on
Windows and macOS).

```bash
pip install -r requirements.txt
```

Then run whichever script you need:

```bash
python attendance_recorder.py      # on each hall's PC, during class
python attendance_analyzer.py      # once, on the merged log file
```

### macOS: "Tkinter" / "macOS 26 (2603) or later required" error

On very new macOS versions (e.g. macOS 26 "Tahoe"), the Tcl/Tk library
bundled with Apple's system Python is too old to recognize the OS version,
and `attendance_recorder.py` will fail to open a window with an error like
`macOS 26 (2603) or later required, have instead 16 (1603)!`. This is an
Apple/Tcl-Tk issue, not a bug in this script. Fix it with a Python build
that has a current Tcl/Tk, e.g. via Homebrew:

```bash
brew install python-tk
/opt/homebrew/bin/python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python attendance_recorder.py
```

### Windows

No known Tkinter issues. Install Python from python.org (Tkinter is included
by default), then:

```bat
pip install -r requirements.txt
python attendance_recorder.py
```
