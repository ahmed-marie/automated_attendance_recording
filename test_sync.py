"""
Automated test harness for the sync layer, simulating two TA laptops on
localhost (different ports/db files stand in for two physical machines on
a shared hotspot). Run: python3 test_sync.py
"""
import os
import sys
import time
import shutil

sys.path.insert(0, os.path.dirname(__file__))
import attendance_db as db
import sync_service as sync

WORKDIR = "/tmp/attendance_sync_test"
DB_A = os.path.join(WORKDIR, "a.db")
DB_B = os.path.join(WORKDIR, "b.db")
PORT_A = 8801
PORT_B = 8802

passed, failed = [], []


def check(label, condition):
    if condition:
        passed.append(label)
        print(f"  PASS  {label}")
    else:
        failed.append(label)
        print(f"  FAIL  {label}")


def reset():
    if os.path.exists(WORKDIR):
        shutil.rmtree(WORKDIR)
    os.makedirs(WORKDIR)
    db.init_db(DB_A)
    db.init_db(DB_B)


# ---------------------------------------------------------------- Test 1 --
def test_basic_bidirectional_sync():
    print("\n[Test 1] Basic bidirectional sync (scans + roster + session close)")
    reset()
    server_a = sync.start_server(DB_A, {"ta_name": "TA_A", "is_controller": True}, port=PORT_A)
    server_b = sync.start_server(DB_B, {"ta_name": "TA_B", "is_controller": False}, port=PORT_B)
    peer_a_to_b = f"http://127.0.0.1:{PORT_B}"
    peer_b_to_a = f"http://127.0.0.1:{PORT_A}"

    db.upsert_roster("UID0001", "40-1234", "Alice", "Engineering", db_path=DB_A)
    db.start_session("20260819_1", db_path=DB_A)
    db.insert_scan("A|UID0001|1", "UID0001", "40-1234", "Alice", "Engineering", "20260819_1", "TA_A", db_path=DB_A)

    db.upsert_roster("UID0002", "40-5678", "Bob", "Business Informatics", db_path=DB_B)
    db.insert_scan("B|UID0002|1", "UID0002", "40-5678", "Bob", "Business Informatics", "20260819_1", "TA_B", db_path=DB_B)

    ok, msg, info = sync.sync_once(peer_a_to_b, DB_A)
    check("A->B sync reports ok", ok)
    check("A learns B's device info", info == {"ta_name": "TA_B", "is_controller": False})

    scans_a = {s["id"] for s in db.get_all_scans(DB_A)}
    scans_b = {s["id"] for s in db.get_all_scans(DB_B)}
    check("A now has both scans", scans_a == {"A|UID0001|1", "B|UID0002|1"})
    check("B has both scans after A pushed (server-side merge)", scans_b == {"A|UID0001|1", "B|UID0002|1"})
    check("A has Bob in roster now", "UID0002" in db.get_roster(DB_A))
    check("B has Alice in roster now", "UID0001" in db.get_roster(DB_B))
    check("B sees the open session started on A", db.get_open_session(DB_B) is not None)

    db.close_session("20260819_1", "Hall 1", "Dr. Karim", 60, db_path=DB_A)
    ok, msg, info = sync.sync_once(peer_a_to_b, DB_A)
    check("Second sync ok", ok)
    check("B now sees the session closed with correct duration",
          db.get_all_sessions(DB_B)["20260819_1"]["duration_minutes"] == 60)
    check("A's open-session query now returns None (closed)", db.get_open_session(DB_A) is None)

    server_a.shutdown()
    server_b.shutdown()


# ---------------------------------------------------------------- Test 2 --
def test_duplicate_scan_idempotent():
    print("\n[Test 2] Re-syncing the same scan twice never duplicates it")
    reset()
    server_b = sync.start_server(DB_B, {"ta_name": "TA_B", "is_controller": False}, port=PORT_B)
    peer_b = f"http://127.0.0.1:{PORT_B}"

    db.start_session("20260819_1", db_path=DB_A)
    db.insert_scan("A|UID0001|1", "UID0001", "40-1234", "Alice", "Engineering", "20260819_1", "TA_A", db_path=DB_A)

    sync.sync_once(peer_b, DB_A)
    sync.sync_once(peer_b, DB_A)  # sync again with nothing new
    sync.sync_once(peer_b, DB_A)  # and again

    check("Only one copy of the scan exists on A", len(db.get_all_scans(DB_A)) == 1)
    check("Only one copy of the scan exists on B", len(db.get_all_scans(DB_B)) == 1)

    server_b.shutdown()


# ---------------------------------------------------------------- Test 3 --
def test_partner_unreachable():
    print("\n[Test 3] Partner unreachable (simulates 'kill WiFi mid-session')")
    reset()
    # No server started for B at all -- A tries to sync and should fail gracefully.
    peer_b = f"http://127.0.0.1:{PORT_B}"
    db.start_session("20260819_1", db_path=DB_A)
    db.insert_scan("A|UID0001|1", "UID0001", "40-1234", "Alice", "Engineering", "20260819_1", "TA_A", db_path=DB_A)

    ok, msg, info = sync.sync_once(peer_b, DB_A, timeout=1)
    check("Sync reports failure (not an exception)", ok is False)
    check("Local data on A is untouched and still valid", len(db.get_all_scans(DB_A)) == 1)
    check("Local scan still readable after failed sync", db.get_all_scans(DB_A)[0]["uid"] == "UID0001")


# ---------------------------------------------------------------- Test 4 --
def test_reconnect_after_outage():
    print("\n[Test 4] Recovery: B is down, A scans 3 students, B comes back, catches up")
    reset()
    peer_b = f"http://127.0.0.1:{PORT_B}"
    db.start_session("20260819_1", db_path=DB_A)
    for i in range(3):
        db.insert_scan(f"A|UID000{i}|1", f"UID000{i}", f"40-{i}", f"Student{i}", "Engineering", "20260819_1", "TA_A", db_path=DB_A)

    ok, _, _ = sync.sync_once(peer_b, DB_A, timeout=1)
    check("Sync fails while B's server is down", ok is False)

    # B "comes back online"
    server_b = sync.start_server(DB_B, {"ta_name": "TA_B", "is_controller": False}, port=PORT_B)
    time.sleep(0.2)
    ok, msg, _ = sync.sync_once(peer_b, DB_A, timeout=3)
    check("Sync succeeds once B is back", ok)
    check("B has all 3 scans it missed while offline", len(db.get_all_scans(DB_B)) == 3)

    server_b.shutdown()


# ---------------------------------------------------------------- Test 5 --
def test_restart_mid_sync():
    print("\n[Test 5] Restarting an app (fresh process/connection) mid-session doesn't lose or duplicate data")
    reset()
    server_b = sync.start_server(DB_B, {"ta_name": "TA_B", "is_controller": False}, port=PORT_B)
    peer_b = f"http://127.0.0.1:{PORT_B}"

    db.start_session("20260819_1", db_path=DB_A)
    db.insert_scan("A|UID0001|1", "UID0001", "40-1234", "Alice", "Engineering", "20260819_1", "TA_A", db_path=DB_A)
    sync.sync_once(peer_b, DB_A)

    # Simulate "app A restarted": a brand new process would just call
    # db.init_db(DB_A) again (idempotent -- CREATE TABLE IF NOT EXISTS) and
    # resume from the same file. We simulate that here directly.
    db.init_db(DB_A)
    check("Data survives a re-init of the same db file", len(db.get_all_scans(DB_A)) == 1)

    # Post-restart, A scans another student and re-syncs.
    db.insert_scan("A|UID0002|1", "UID0002", "40-5678", "Bob", "Business Informatics", "20260819_1", "TA_A", db_path=DB_A)
    ok, msg, _ = sync.sync_once(peer_b, DB_A)
    check("Sync after restart succeeds", ok)
    check("B has both scans, no duplicates, no losses", len(db.get_all_scans(DB_B)) == 2)

    server_b.shutdown()


# ---------------------------------------------------------------- Test 6 --
def test_hall_conflict_detectable_after_merge():
    print("\n[Test 6] A student scanned under two different halls for the same session_id is still detectable downstream")
    reset()
    server_b = sync.start_server(DB_B, {"ta_name": "TA_B", "is_controller": False}, port=PORT_B)
    peer_b = f"http://127.0.0.1:{PORT_B}"

    db.start_session("20260819_1", db_path=DB_A)
    db.insert_scan("A|UID0001|1", "UID0001", "40-1234", "Alice", "Engineering", "20260819_1", "TA_A", db_path=DB_A)
    db.close_session("20260819_1", "Hall 1", "Dr. Karim", 60, db_path=DB_A)

    db.start_session("20260819_1", db_path=DB_B)  # same id, different physical hall (a real edge case)
    db.insert_scan("B|UID0001|1", "UID0001", "40-1234", "Alice", "Engineering", "20260819_1", "TA_B", db_path=DB_B)
    db.close_session("20260819_1", "Hall 2", "Dr. Nourhan", 60, db_path=DB_B)

    sync.sync_once(peer_b, DB_A)
    scans = [s for s in db.get_all_scans(DB_A) if s["uid"] == "UID0001"]
    check("Both scans present after merge (hall-conflict data intact for the analyzer to catch)", len(scans) == 2)

    server_b.shutdown()


if __name__ == "__main__":
    test_basic_bidirectional_sync()
    test_duplicate_scan_idempotent()
    test_partner_unreachable()
    test_reconnect_after_outage()
    test_restart_mid_sync()
    test_hall_conflict_detectable_after_merge()

    print(f"\n{'='*50}\n{len(passed)} passed, {len(failed)} failed\n{'='*50}")
    if failed:
        print("FAILED:", failed)
        sys.exit(1)