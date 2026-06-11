import sqlite3

db_path = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\mpro_reconciliation.db"
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Get all cases
cursor.execute("SELECT id, name, updated_at FROM reconciliation_cases;")
cases = cursor.fetchall()

tables = [
    "po_records",
    "agency_invoice_records",
    "third_party_invoice_records",
    "third_party_monitoring_records"
]

for c in cases:
    print(f"\nCase: {c['name']} (ID: {c['id']})")
    for t in tables:
        try:
            cursor.execute(f"SELECT COUNT(*) FROM {t} WHERE case_id = ?;", (c['id'],))
            count = cursor.fetchone()[0]
            print(f"  - {t}: {count} rows")
        except Exception as e:
            print(f"  - Error counting {t}: {e}")

conn.close()
