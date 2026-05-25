"""
Payment & Order Models — مذكرتي Pro v18
Handles orders, payment receipts, download codes, and audit logs.
Pure data layer — no Flask dependencies.
"""
from __future__ import annotations
import hashlib
import json
import os
import secrets
import sqlite3
import time
from dataclasses import dataclass, field
from typing import Optional

DB_PATH = os.environ.get("DB_PATH", "mathkarati_payments.db")

# ─── Schema ──────────────────────────────────────────────────────────────────

SCHEMA_SQL = """
PRAGMA journal_mode=WAL;
PRAGMA foreign_keys=ON;

CREATE TABLE IF NOT EXISTS orders (
    id            TEXT PRIMARY KEY,
    presentation_id TEXT NOT NULL,
    student_name  TEXT NOT NULL,
    email         TEXT NOT NULL DEFAULT '',
    phone         TEXT NOT NULL DEFAULT '',
    degree        TEXT NOT NULL DEFAULT 'licence',   -- licence / master / doctorat
    title_ar      TEXT NOT NULL DEFAULT '',
    amount        INTEGER NOT NULL DEFAULT 800,
    currency      TEXT NOT NULL DEFAULT 'DZD',
    status        TEXT NOT NULL DEFAULT 'pending',   -- pending / approved / rejected / downloaded
    payment_method TEXT NOT NULL DEFAULT '',         -- ccp / baridimob
    receipt_path  TEXT,                              -- path to uploaded receipt file
    receipt_mime  TEXT,
    download_code TEXT UNIQUE,
    code_used     INTEGER NOT NULL DEFAULT 0,
    code_expires  REAL,                              -- Unix timestamp
    created_at    REAL NOT NULL,
    updated_at    REAL NOT NULL,
    approved_at   REAL,
    downloaded_at REAL,
    download_ip   TEXT,
    admin_note    TEXT,
    pptx_path     TEXT                               -- path to stored PPTX (protected)
);

CREATE TABLE IF NOT EXISTS audit_log (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    order_id   TEXT NOT NULL,
    action     TEXT NOT NULL,
    actor      TEXT NOT NULL DEFAULT 'system',     -- admin / student / system
    detail     TEXT,
    ip         TEXT,
    ts         REAL NOT NULL
);

CREATE TABLE IF NOT EXISTS admin_sessions (
    token      TEXT PRIMARY KEY,
    created_at REAL NOT NULL,
    expires_at REAL NOT NULL,
    ip         TEXT
);

CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_code ON orders(download_code);
CREATE INDEX IF NOT EXISTS idx_audit_order ON audit_log(order_id);
"""


def get_db() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH, timeout=10)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    conn = get_db()
    conn.executescript(SCHEMA_SQL)
    conn.commit()
    conn.close()


# ─── Order helpers ────────────────────────────────────────────────────────────

def new_order_id() -> str:
    return "ORD-" + secrets.token_hex(6).upper()


def new_download_code() -> str:
    """8-char alphanumeric code, uppercase, easy to type."""
    alphabet = "ABCDEFGHJKLMNPQRSTUVWXYZ23456789"  # no I/O/1/0 ambiguity
    return "".join(secrets.choice(alphabet) for _ in range(8))


def create_order(
    presentation_id: str,
    student_name: str,
    email: str,
    phone: str,
    degree: str,
    title_ar: str,
    payment_method: str,
) -> dict:
    now = time.time()
    order_id = new_order_id()
    conn = get_db()
    try:
        conn.execute(
            """INSERT INTO orders
               (id, presentation_id, student_name, email, phone, degree, title_ar,
                payment_method, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?)""",
            (order_id, presentation_id, student_name, email, phone, degree, title_ar,
             payment_method, now, now),
        )
        conn.commit()
        _log(conn, order_id, "created", "student", f"method={payment_method}")
        conn.commit()
    finally:
        conn.close()
    return get_order(order_id)


def get_order(order_id: str) -> Optional[dict]:
    conn = get_db()
    try:
        row = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        return dict(row) if row else None
    finally:
        conn.close()


def list_orders(status: str = None, limit: int = 200) -> list[dict]:
    conn = get_db()
    try:
        if status:
            rows = conn.execute(
                "SELECT * FROM orders WHERE status=? ORDER BY created_at DESC LIMIT ?",
                (status, limit),
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM orders ORDER BY created_at DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def attach_receipt(order_id: str, path: str, mime: str):
    conn = get_db()
    try:
        conn.execute(
            "UPDATE orders SET receipt_path=?, receipt_mime=?, updated_at=? WHERE id=?",
            (path, mime, time.time(), order_id),
        )
        conn.commit()
        _log(conn, order_id, "receipt_uploaded", "student", f"mime={mime}")
        conn.commit()
    finally:
        conn.close()


def approve_order(order_id: str, admin_note: str = "", code_ttl_hours: int = 48) -> str:
    """Generate download code and mark approved."""
    code = new_download_code()
    expires = time.time() + code_ttl_hours * 3600
    now = time.time()
    conn = get_db()
    try:
        conn.execute(
            """UPDATE orders
               SET status='approved', download_code=?, code_expires=?,
                   approved_at=?, updated_at=?, admin_note=?
               WHERE id=?""",
            (code, expires, now, now, admin_note, order_id),
        )
        conn.commit()
        _log(conn, order_id, "approved", "admin", f"code_ttl={code_ttl_hours}h")
        conn.commit()
    finally:
        conn.close()
    return code


def reject_order(order_id: str, admin_note: str = ""):
    now = time.time()
    conn = get_db()
    try:
        conn.execute(
            "UPDATE orders SET status='rejected', updated_at=?, admin_note=? WHERE id=?",
            (now, admin_note, order_id),
        )
        conn.commit()
        _log(conn, order_id, "rejected", "admin", admin_note)
        conn.commit()
    finally:
        conn.close()


def store_pptx(order_id: str, path: str):
    conn = get_db()
    try:
        conn.execute(
            "UPDATE orders SET pptx_path=?, updated_at=? WHERE id=?",
            (path, time.time(), order_id),
        )
        conn.commit()
    finally:
        conn.close()


def redeem_code(code: str, ip: str) -> Optional[dict]:
    """
    Validate code → mark used → return order dict, or None if invalid.
    Thread-safe via SQLite row-level update.
    """
    now = time.time()
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT * FROM orders WHERE download_code=?", (code,)
        ).fetchone()
        if not row:
            return None
        order = dict(row)
        if order["status"] != "approved":
            return None
        if order["code_used"]:
            return None
        if order["code_expires"] and now > order["code_expires"]:
            return None
        # Mark used
        conn.execute(
            """UPDATE orders
               SET code_used=1, status='downloaded', downloaded_at=?, download_ip=?, updated_at=?
               WHERE id=? AND code_used=0""",
            (now, ip, now, order["id"]),
        )
        if conn.execute("SELECT changes()").fetchone()[0] == 0:
            return None  # race condition
        conn.commit()
        _log(conn, order["id"], "downloaded", "student", f"ip={ip}")
        conn.commit()
        return get_order(order["id"])
    finally:
        conn.close()


def get_stats() -> dict:
    conn = get_db()
    try:
        total = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
        by_status = {
            r["status"]: r["cnt"]
            for r in conn.execute(
                "SELECT status, COUNT(*) as cnt FROM orders GROUP BY status"
            ).fetchall()
        }
        revenue = conn.execute(
            "SELECT COALESCE(SUM(amount),0) FROM orders WHERE status IN ('approved','downloaded')"
        ).fetchone()[0]
        today_start = time.time() - (time.time() % 86400)
        today = conn.execute(
            "SELECT COUNT(*) FROM orders WHERE created_at >= ?", (today_start,)
        ).fetchone()[0]
        return {
            "total": total,
            "by_status": by_status,
            "revenue_dzd": revenue,
            "today": today,
            "pending": by_status.get("pending", 0),
            "approved": by_status.get("approved", 0),
            "downloaded": by_status.get("downloaded", 0),
            "rejected": by_status.get("rejected", 0),
        }
    finally:
        conn.close()


def _log(conn: sqlite3.Connection, order_id: str, action: str, actor: str, detail: str = ""):
    conn.execute(
        "INSERT INTO audit_log (order_id, action, actor, detail, ts) VALUES (?,?,?,?,?)",
        (order_id, action, actor, detail, time.time()),
    )


def get_audit(order_id: str) -> list[dict]:
    conn = get_db()
    try:
        rows = conn.execute(
            "SELECT * FROM audit_log WHERE order_id=? ORDER BY ts DESC", (order_id,)
        ).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ─── Admin session ────────────────────────────────────────────────────────────

ADMIN_PASSWORD_HASH = os.environ.get(
    "ADMIN_PASSWORD_HASH",
    "9bf7847c3f1d260cbafe797fd078f34ecce9bbdc5121208452e7528fc73732b7",
)


def admin_login(password: str, ip: str) -> Optional[str]:
    hashed = hashlib.sha256(password.encode()).hexdigest()
    if not secrets.compare_digest(hashed, ADMIN_PASSWORD_HASH):
        return None
    token = secrets.token_hex(32)
    now = time.time()
    conn = get_db()
    try:
        conn.execute(
            "INSERT INTO admin_sessions (token, created_at, expires_at, ip) VALUES (?,?,?,?)",
            (token, now, now + 8 * 3600, ip),
        )
        conn.commit()
    finally:
        conn.close()
    return token


def admin_check(token: str) -> bool:
    if not token:
        return False
    conn = get_db()
    try:
        row = conn.execute(
            "SELECT expires_at FROM admin_sessions WHERE token=?", (token,)
        ).fetchone()
        return bool(row and row["expires_at"] > time.time())
    finally:
        conn.close()
