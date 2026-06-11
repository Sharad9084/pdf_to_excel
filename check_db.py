import sqlite3
import os

db_path = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\mpro_reconciliation.db"
print(f"Checking database at: {db_path}")
print(f"File exists: {os.path.exists(db_path)}")
if os.path.exists(db_path):
    print(f"File size: {os.path.getsize(db_path)} bytes")

conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row

# Get tables
cursor = conn.cursor()
cursor.execute("SELECT name FROM sqlite_master WHERE type='table';")
tables = [row[0] for row in cursor.fetchall()]
print(f"\nTables in database: {tables}")

# Check cases
try:
    cursor.execute("SELECT id, name, updated_at FROM reconciliation_cases;")
    cases = cursor.fetchall()
    print(f"\nCases ({len(cases)}):")
    for c in cases:
        print(f"  - ID: {c['id']}, Name: {c['name']}, Updated: {c['updated_at']}")
except Exception as e:
    print(f"Error fetching cases: {e}")

# Check row counts in tables
for t in tables:
    try:
        cursor.execute(f"SELECT COUNT(*) FROM {t};")
        count = cursor.fetchone()[0]
        print(f"Table '{t}': {count} rows")
    except Exception as e:
        print(f"Error counting table {t}: {e}")

conn.close()
