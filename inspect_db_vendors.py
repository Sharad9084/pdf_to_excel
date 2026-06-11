import sqlite3
import json

db_path = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\mpro_reconciliation.db"
conn = sqlite3.connect(db_path)
conn.row_factory = sqlite3.Row
cursor = conn.cursor()

case_id = "e3ac65e8-3ab1-4bcf-80c2-c6603ab87716"

print(f"=== ANALYZING SQLite RECORDS FOR CASE {case_id} ===")

# Query rows for this case
cursor.execute("SELECT raw_json FROM third_party_invoice_records WHERE case_id = ?;", (case_id,))
rows = [json.loads(row[0]) for row in cursor.fetchall()]

print(f"Total records in SQLite for this case: {len(rows)}")

# Count unique files
files = set()
channels = {}
vendors = {}

for r in rows:
    file_name = r.get("File Name") or r.get("PDF File Name") or "Unknown File"
    files.add(file_name)
    
    channel = r.get("Channel Name") or r.get("Channel") or "Unknown Channel"
    channels[channel] = channels.get(channel, 0) + 1
    
    vendor = r.get("Third Party Vendor Name") or r.get("Broadcaster Name") or r.get("Broadcaster_Name") or "Unknown Vendor"
    vendors[vendor] = vendors.get(vendor, 0) + 1

print(f"\nTotal unique PDF files processed: {len(files)}")
print(f"Unique Channels found in rows: {list(channels.keys())}")
print("\nChannel distribution in SQLite rows:")
for chan, count in channels.items():
    print(f"  - {chan}: {count} rows")

print("\nVendor distribution in SQLite rows (top 15):")
sorted_vendors = sorted(vendors.items(), key=lambda x: x[1], reverse=True)
for ven, count in sorted_vendors[:15]:
    print(f"  - {ven}: {count} rows")

conn.close()
