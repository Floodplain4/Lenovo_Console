import csv
import sqlite3
from pathlib import Path

DB_FILE = Path("lenovo_tracker.db")
CSV_FILE = Path("lcd_log.csv")

print("Current folder:", Path.cwd())
print("CSV exists:", CSV_FILE.exists(), CSV_FILE.resolve())
print("DB path:", DB_FILE.resolve())

conn = sqlite3.connect(DB_FILE)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS cases (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    serial_number TEXT,
    work_order TEXT,
    case_number TEXT,
    status TEXT,
    parts TEXT,
    notes TEXT,
    created_at TEXT,
    updated_at TEXT
)
""")

with open(CSV_FILE, newline="", encoding="utf-8-sig") as f:
    reader = csv.DictReader(f)

    print("CSV headers found:")
    print(reader.fieldnames)

    count = 0

    for row in reader:
        print("First row sample:" if count == 0 else "", row if count == 0 else "")

        cursor.execute("""
            INSERT INTO cases (
                serial_number,
                work_order,
                case_number,
                status,
                parts,
                notes,
                created_at,
                updated_at
            )
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            row.get("Serial Number", ""),
            row.get("Work Order", ""),
            row.get("Case Number", ""),
            row.get("Status", ""),
            row.get("Parts", ""),
            row.get("Notes", ""),
            row.get("Created At", ""),
            row.get("Updated At", "")
        ))

        count += 1

conn.commit()

cursor.execute("SELECT COUNT(*) FROM cases")
db_count = cursor.fetchone()[0]

conn.close()

print(f"Rows imported from CSV: {count}")
print(f"Rows now in database: {db_count}")
print("Import complete.")