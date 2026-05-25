"""
Database Layer — مذكرتي Pro v18 Commercial
SQLite with full order lifecycle, download codes, and analytics.
"""
from __future__ import annotations
import sqlite3
import uuid
import secrets
import string
import hashlib
import os
import json
from datetime import datetime, timedelta
from contextlib import contextmanager
from dataclasses import dataclass, field
from typing import Optional

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "data.db")

SCHEMA = """
CREATE TABLE IF NOT EXISTS orders (
    id TEXT PRIMARY KEY,
    preview_token TEXT UNIQUE NOT NULL,
    pptx_path TEXT,
    pdf_path TEXT,
    student_name TEXT NOT NULL,
    title_ar TEXT NOT NULL,
    engine TEXT DEFAULT 'canva',
    theme TEXT DEFAULT 'navy_gold',
    slide_count INTEGER DEFAULT 0,
    request_json TEXT,
    status TEXT DEFAULT 'preview',
    receipt_path TEXT,
    receipt_filename TEXT,
    admin_note TEXT,
    created_at TEXT NOT NULL,
    submitted_at TEXT,
    approved_at TEXT,
    rejected_at TEXT
);

CREATE TABLE IF NOT EXISTS download_codes (
    code TEXT PRIMARY KEY,
    order_id TEXT NOT NULL,
    max_uses INTEGER DEFAULT 1,
    uses_count INTEGER DEFAULT 0,
    expires_at TEXT,
    is_active INTEGER DEFAULT 1,
    created_at TEXT NOT NULL,
    last_used_at TEXT,
    last_used_ip TEXT,
    FOREIGN KEY(order_id) REFERENCES orders(id)
);

CREATE TABLE IF NOT EXISTS download_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    code TEXT NOT NULL,
    order_id TEXT NOT NULL,
    ip TEXT,
    user_agent TEXT,
    downloaded_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_orders_status ON orders(status);
CREATE INDEX IF NOT EXISTS idx_orders_preview_token ON orders(preview_token);
CREATE INDEX IF NOT EXISTS idx_codes_order_id ON download_codes(order_id);
"""

DEFAULT_SETTINGS = {
    "price": "800",
    "currency": "دج",
    "payment_ccp": "",
    "payment_baridimob": "",
    "payment_name": "",
    "payment_info": "",
    "admin_password": "admin1234",
    "site_name": "مذكرتي Pro",
    "contact_email": "",
    "code_validity_days": "7",
}


@contextmanager
def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    try:
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def init_db():
    with get_db() as conn:
        conn.executescript(SCHEMA)
        # Insert default settings if missing
        for key, value in DEFAULT_SETTINGS.items():
            conn.execute(
                "INSERT OR IGNORE INTO settings(key, value, updated_at) VALUES(?,?,?)",
                (key, value, _now())
            )


def _now() -> str:
    return datetime.utcnow().isoformat()


def _gen_token() -> str:
    return secrets.token_urlsafe(32)


def _gen_code() -> str:
    chars = string.ascii_uppercase + string.digits
    chars = chars.replace("O", "").replace("0", "").replace("I", "").replace("1", "")
    return "".join(secrets.choice(chars) for _ in range(8))


# ─────────────────────────────────────────────────
# ORDERS
# ─────────────────────────────────────────────────

def create_order(student_name: str, title_ar: str, engine: str, theme: str,
                 request_json: dict, slide_count: int) -> dict:
    oid = str(uuid.uuid4())
    token = _gen_token()
    now = _now()
    with get_db() as conn:
        conn.execute(
            """INSERT INTO orders(id, preview_token, student_name, title_ar, engine, theme,
               slide_count, request_json, status, created_at)
               VALUES(?,?,?,?,?,?,?,?,?,?)""",
            (oid, token, student_name, title_ar, engine, theme,
             slide_count, json.dumps(request_json, ensure_ascii=False), "preview", now)
        )
    return {"id": oid, "preview_token": token}


def set_order_files(order_id: str, pptx_path: str, pdf_path: str):
    with get_db() as conn:
        conn.execute(
            "UPDATE orders SET pptx_path=?, pdf_path=? WHERE id=?",
            (pptx_path, pdf_path, order_id)
        )


def submit_receipt(order_id: str, receipt_path: str, receipt_filename: str):
    with get_db() as conn:
        conn.execute(
            "UPDATE orders SET status='pending', receipt_path=?, receipt_filename=?, submitted_at=? WHERE id=?",
            (receipt_path, receipt_filename, _now(), order_id)
        )


def get_order(order_id: str) -> Optional[dict]:
    with get_db() as conn:
        row = conn.execute("SELECT * FROM orders WHERE id=?", (order_id,)).fetchone()
        return dict(row) if row else None


def get_order_by_token(token: str) -> Optional[dict]:
    with get_db() as conn:
        row = conn.execute("SELECT * FROM orders WHERE preview_token=?", (token,)).fetchone()
        return dict(row) if row else None


def list_orders(status: str = None, limit: int = 100) -> list:
    with get_db() as conn:
        if status:
            rows = conn.execute(
                "SELECT * FROM orders WHERE status=? ORDER BY created_at DESC LIMIT ?",
                (status, limit)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM orders ORDER BY created_at DESC LIMIT ?", (limit,)
            ).fetchall()
        return [dict(r) for r in rows]


def approve_order(order_id: str, note: str = "") -> str:
    """Approve order and create download code. Returns the code."""
    code = _gen_code()
    settings = get_settings()
    days = int(settings.get("code_validity_days", 7))
    expires = (datetime.utcnow() + timedelta(days=days)).isoformat()
    now = _now()
    with get_db() as conn:
        conn.execute(
            "UPDATE orders SET status='approved', approved_at=?, admin_note=? WHERE id=?",
            (now, note, order_id)
        )
        conn.execute(
            """INSERT INTO download_codes(code, order_id, max_uses, uses_count,
               expires_at, is_active, created_at)
               VALUES(?,?,1,0,?,1,?)""",
            (code, order_id, expires, now)
        )
    return code


def reject_order(order_id: str, note: str = ""):
    with get_db() as conn:
        conn.execute(
            "UPDATE orders SET status='rejected', rejected_at=?, admin_note=? WHERE id=?",
            (_now(), note, order_id)
        )


# ─────────────────────────────────────────────────
# DOWNLOAD CODES
# ─────────────────────────────────────────────────

def validate_code(code: str, ip: str = "", user_agent: str = "") -> dict:
    """Returns {'valid': bool, 'order_id': str, 'error': str}"""
    code = code.strip().upper()
    with get_db() as conn:
        row = conn.execute(
            "SELECT * FROM download_codes WHERE code=?", (code,)
        ).fetchone()

        if not row:
            return {"valid": False, "error": "الكود غير صحيح"}
        row = dict(row)

        if not row["is_active"]:
            return {"valid": False, "error": "الكود معطّل"}

        if row["expires_at"] and row["expires_at"] < _now():
            return {"valid": False, "error": "انتهت صلاحية الكود"}

        if row["uses_count"] >= row["max_uses"]:
            return {"valid": False, "error": "تم استخدام الكود مسبقاً"}

        # Record usage
        now = _now()
        conn.execute(
            "UPDATE download_codes SET uses_count=uses_count+1, last_used_at=?, last_used_ip=? WHERE code=?",
            (now, ip, code)
        )
        conn.execute(
            "INSERT INTO download_logs(code, order_id, ip, user_agent, downloaded_at) VALUES(?,?,?,?,?)",
            (code, row["order_id"], ip, user_agent, now)
        )
        return {"valid": True, "order_id": row["order_id"], "error": ""}


def list_codes(order_id: str = None) -> list:
    with get_db() as conn:
        if order_id:
            rows = conn.execute(
                "SELECT * FROM download_codes WHERE order_id=? ORDER BY created_at DESC",
                (order_id,)
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT * FROM download_codes ORDER BY created_at DESC LIMIT 200"
            ).fetchall()
        return [dict(r) for r in rows]


def deactivate_code(code: str):
    with get_db() as conn:
        conn.execute("UPDATE download_codes SET is_active=0 WHERE code=?", (code,))


def create_manual_code(order_id: str, max_uses: int = 1, days: int = 7) -> str:
    code = _gen_code()
    expires = (datetime.utcnow() + timedelta(days=days)).isoformat()
    with get_db() as conn:
        conn.execute(
            """INSERT INTO download_codes(code, order_id, max_uses, uses_count,
               expires_at, is_active, created_at) VALUES(?,?,?,0,?,1,?)""",
            (code, order_id, max_uses, expires, _now())
        )
    return code


# ─────────────────────────────────────────────────
# SETTINGS
# ─────────────────────────────────────────────────

def get_settings() -> dict:
    with get_db() as conn:
        rows = conn.execute("SELECT key, value FROM settings").fetchall()
        return {r["key"]: r["value"] for r in rows}


def update_settings(updates: dict):
    now = _now()
    with get_db() as conn:
        for key, value in updates.items():
            conn.execute(
                "INSERT OR REPLACE INTO settings(key, value, updated_at) VALUES(?,?,?)",
                (key, str(value), now)
            )


def check_admin_password(password: str) -> bool:
    settings = get_settings()
    return password == settings.get("admin_password", "admin1234")


# ─────────────────────────────────────────────────
# ANALYTICS
# ─────────────────────────────────────────────────

def get_stats() -> dict:
    with get_db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM orders").fetchone()[0]
        pending = conn.execute("SELECT COUNT(*) FROM orders WHERE status='pending'").fetchone()[0]
        approved = conn.execute("SELECT COUNT(*) FROM orders WHERE status='approved'").fetchone()[0]
        rejected = conn.execute("SELECT COUNT(*) FROM orders WHERE status='rejected'").fetchone()[0]
        downloads = conn.execute("SELECT COUNT(*) FROM download_logs").fetchone()[0]

        # Revenue
        settings = get_settings()
        price = int(settings.get("price", 800))
        revenue = approved * price

        # Engine breakdown
        engines = conn.execute(
            "SELECT engine, COUNT(*) as cnt FROM orders GROUP BY engine ORDER BY cnt DESC"
        ).fetchall()

        # Theme breakdown
        themes = conn.execute(
            "SELECT theme, COUNT(*) as cnt FROM orders GROUP BY theme ORDER BY cnt DESC LIMIT 5"
        ).fetchall()

        # Last 7 days
        week_ago = (datetime.utcnow() - timedelta(days=7)).isoformat()
        recent = conn.execute(
            "SELECT COUNT(*) FROM orders WHERE created_at > ?", (week_ago,)
        ).fetchone()[0]

        return {
            "total": total,
            "pending": pending,
            "approved": approved,
            "rejected": rejected,
            "preview": total - pending - approved - rejected,
            "downloads": downloads,
            "revenue": revenue,
            "revenue_formatted": f"{revenue:,} {settings.get('currency', 'دج')}",
            "engines": [{"engine": r["engine"], "count": r["cnt"]} for r in engines],
            "themes": [{"theme": r["theme"], "count": r["cnt"]} for r in themes],
            "recent_7days": recent,
        }
