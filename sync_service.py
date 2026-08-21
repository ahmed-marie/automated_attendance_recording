"""
sync_service.py
=================
Peer-to-peer sync between the two TA laptops over a local network (e.g. a
shared mobile hotspot) -- no internet required, no external service.

Each laptop runs a small built-in HTTP server (Python's standard library
only) exposing:
    GET  /ping   -- lightweight reachability + identity check
    POST /sync   -- caller sends its full local dataset; the peer merges it
                     in and replies with its OWN full dataset, which the
                     caller then merges back. One request/response = a
                     complete two-way merge.

Data volumes here are small (a few hundred students, a few thousand scans
across a whole term), so exchanging full tables every round is simple and
fast -- there are no incremental cursors to get wrong, which matters given
this needs to be reliable, not just fast.
"""

import json
import socket
import threading
import time
import urllib.request
import urllib.error
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

import attendance_db as db

DEFAULT_PORT = 8765
STALE_AFTER_SECONDS = 30  # how long since last activity before the GUI treats sync as "stale", not "live"


class SyncStatus:
    """Sync activity is bidirectional -- either laptop can be the one that
    initiates a round -- so status has to be a single object that BOTH the
    outbound sync loop and the inbound HTTP handler update, not an
    attribute that only the initiating side ever touches. Thread-safe:
    written from the background sync thread and from per-request server
    threads, read from the GUI's main-thread poll."""

    def __init__(self):
        self._lock = threading.Lock()
        self.text = "Not connected"
        self.last_activity_ts = None
        self.last_remote_device_info = None

    def set(self, text, remote_device_info=None, mark_activity=False):
        with self._lock:
            self.text = text
            if remote_device_info is not None:
                self.last_remote_device_info = remote_device_info
            if mark_activity:
                self.last_activity_ts = time.time()

    def get(self):
        with self._lock:
            return self.text, self.last_activity_ts, self.last_remote_device_info

    def is_live(self):
        _, ts, _ = self.get()
        return ts is not None and (time.time() - ts) < STALE_AFTER_SECONDS


def get_local_ip():
    """Best-effort local network IP. Uses a UDP 'connect' (no packets sent,
    just asks the OS routing table which interface would be used) so it
    works even with zero internet access -- exactly the hotspot scenario."""
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    try:
        s.connect(("192.168.137.1", 80))  # typical Windows Mobile Hotspot gateway; doesn't need to respond
        ip = s.getsockname()[0]
    except OSError:
        ip = "127.0.0.1"
    finally:
        s.close()
    return ip


def _make_handler(db_path, device_info, status):
    class SyncHandler(BaseHTTPRequestHandler):
        def log_message(self, fmt, *args):
            pass  # keep the TA's window quiet

        def _send_json(self, obj, http_status=200):
            body = json.dumps(obj).encode("utf-8")
            self.send_response(http_status)
            self.send_header("Content-Type", "application/json")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        def do_GET(self):
            if self.path == "/ping":
                self._send_json({"ok": True, "device_info": device_info})
            else:
                self._send_json({"ok": False, "error": "not found"}, http_status=404)

        def do_POST(self):
            if self.path != "/sync":
                self._send_json({"ok": False, "error": "not found"}, http_status=404)
                return
            try:
                length = int(self.headers.get("Content-Length", 0))
                payload = json.loads(self.rfile.read(length).decode("utf-8"))
                remote_info = payload.get("device_info")
                added = db.merge_payload(payload, db_path=db_path)
                my_data = db.export_all(db_path=db_path)
                stamp = time.strftime("%H:%M:%S")
                who = (remote_info or {}).get("ta_name", "partner")
                status.set(
                    f"[{stamp}] Received sync from {who} ({added} new scan(s))",
                    remote_device_info=remote_info,
                    mark_activity=True,
                )
                self._send_json({"ok": True, "added_scans": added, "device_info": device_info, **my_data})
            except Exception as e:
                self._send_json({"ok": False, "error": str(e)}, http_status=500)

    return SyncHandler


def start_server(db_path, device_info, status, port=DEFAULT_PORT):
    """Starts the inbound sync server in a daemon thread. Returns the server
    object (call .shutdown() to stop it, mainly useful for tests)."""
    handler = _make_handler(db_path, device_info, status)
    server = ThreadingHTTPServer(("0.0.0.0", port), handler)
    thread = threading.Thread(target=server.serve_forever, daemon=True)
    thread.start()
    return server


def _http_json(method, url, payload=None, timeout=4):
    data = json.dumps(payload).encode("utf-8") if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=timeout) as resp:
        return json.loads(resp.read().decode("utf-8"))


def ping(peer_base_url, timeout=3):
    return _http_json("GET", f"{peer_base_url}/ping", timeout=timeout)


def sync_once(peer_base_url, db_path, my_device_info=None, timeout=5):
    """One full bidirectional exchange. Returns (ok, message, remote_device_info)."""
    my_data = db.export_all(db_path=db_path)
    if my_device_info:
        my_data["device_info"] = my_device_info
    try:
        remote = _http_json("POST", f"{peer_base_url}/sync", payload=my_data, timeout=timeout)
    except (urllib.error.URLError, socket.timeout, TimeoutError, ConnectionError, OSError) as e:
        return False, f"Could not reach partner: {e}", None
    if not remote.get("ok"):
        return False, f"Partner reported an error: {remote.get('error')}", None
    added = db.merge_payload(remote, db_path=db_path)
    return True, f"Synced ({added} new scan(s) pulled)", remote.get("device_info")


class SyncManager:
    """Runs sync_once() on a repeating background timer once started.
    Only touches the DB via short-lived connections (see attendance_db.py)
    and only touches plain attributes / the shared SyncStatus here -- never
    Tk widgets directly, since this runs on a background thread. The GUI
    polls SyncStatus from its own main-thread timer instead."""

    def __init__(self, db_path, status, device_info=None):
        self.db_path = db_path
        self.status = status
        self.device_info = device_info
        self.peer_base_url = None
        self._stop_event = threading.Event()
        self._thread = None

    def start(self, peer_base_url, interval_seconds=12):
        self.peer_base_url = peer_base_url
        self._stop_event.clear()
        if self._thread and self._thread.is_alive():
            return
        self._thread = threading.Thread(target=self._loop, args=(interval_seconds,), daemon=True)
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        self.peer_base_url = None

    def sync_now(self):
        if not self.peer_base_url:
            self.status.set("Not connected to a partner yet.")
            return False
        ok, msg, remote_info = sync_once(self.peer_base_url, self.db_path, my_device_info=self.device_info)
        stamp = time.strftime("%H:%M:%S")
        self.status.set(f"[{stamp}] {msg}", remote_device_info=remote_info, mark_activity=ok)
        return ok

    def _loop(self, interval_seconds):
        while not self._stop_event.is_set():
            self.sync_now()
            self._stop_event.wait(interval_seconds)