import sqlite3
from pathlib import Path

DB_FILE = Path(__file__).with_name("lenovo_tracker.db")

def get_connection():
    return sqlite3.connect(DB_FILE)

def get_all_cases():
    conn = get_connection()
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM cases
        ORDER BY updated_at DESC
    """)

    rows = [dict(row) for row in cursor.fetchall()]
    conn.close()
    return rows