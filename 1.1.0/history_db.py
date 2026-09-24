"""Small SQLite audit/history store for Allison's PDF Extractor."""

import json
import os
import sqlite3
from datetime import datetime


def _db_path():
    root = os.environ.get("LOCALAPPDATA") or os.path.expanduser("~")
    folder = os.path.join(root, "AllisonsPDFExtractor")
    os.makedirs(folder, exist_ok=True)
    return os.path.join(folder, "history.db")


def _connect():
    conn = sqlite3.connect(_db_path())
    conn.execute(
        """
        CREATE TABLE IF NOT EXISTS history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at TEXT NOT NULL,
            action TEXT NOT NULL,
            source_file TEXT,
            template TEXT,
            sales_order TEXT,
            status TEXT,
            output_file TEXT,
            details TEXT
        )
        """
    )
    conn.commit()
    return conn


def log_event(action, source_file="", template="", sales_order="", status="", output_file="", details=None):
    with _connect() as conn:
        conn.execute(
            "INSERT INTO history (created_at, action, source_file, template, sales_order, status, output_file, details) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (
                datetime.now().isoformat(timespec="seconds"), action, source_file, template,
                str(sales_order or ""), status, output_file,
                json.dumps(details or {}, ensure_ascii=False),
            ),
        )


def recent_events(limit=200):
    with _connect() as conn:
        rows = conn.execute(
            "SELECT created_at, action, source_file, template, sales_order, status, output_file, details FROM history ORDER BY id DESC LIMIT ?",
            (int(limit),),
        ).fetchall()
    return rows


def clear_history():
    """Delete all processing-history rows. Returns the number of rows removed."""
    with _connect() as conn:
        count = conn.execute("SELECT COUNT(*) FROM history").fetchone()[0]
        conn.execute("DELETE FROM history")
        # Reset AUTOINCREMENT so a fresh history starts cleanly.
        conn.execute("DELETE FROM sqlite_sequence WHERE name = 'history'")
        conn.commit()
    return int(count)
