import sqlite3

db_path = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\mpro_reconciliation.db"
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

# Get total rows
cursor.execute("SELECT COUNT(*) FROM third_party_invoice_records;")
total = cursor.fetchone()[0]
print(f"Total rows in third_party_invoice_records: {total}")

# Count non-null columns
columns = [
    "third_party_vendor_name",
    "channel_name",
    "invoice_number",
    "invoice_date",
    "po_number",
    "rate_inr",
    "calculated_amount_inr"
]

print("\nNon-null column counts:")
for col in columns:
    cursor.execute(f"SELECT COUNT(*) FROM third_party_invoice_records WHERE {col} IS NOT NULL AND {col} != '';")
    non_null = cursor.fetchone()[0]
    print(f"  - {col}: {non_null} rows")

# Distribution of third_party_vendor_name in db columns
print("\nUnique values of 'third_party_vendor_name' stored in DB column:")
cursor.execute("SELECT third_party_vendor_name, COUNT(*) FROM third_party_invoice_records GROUP BY third_party_vendor_name;")
for row in cursor.fetchall():
    print(f"  - {row[0]}: {row[1]} rows")

# Inspect a raw_json from the table
print("\nKeys inside raw_json for first row:")
cursor.execute("SELECT raw_json FROM third_party_invoice_records LIMIT 1;")
raw = cursor.fetchone()[0]
raw_dict = json = sqlite3.dbapi2.json.loads(raw)
for k, v in list(raw_dict.items())[:15]:
    print(f"  {k}: {v}")

conn.close()
