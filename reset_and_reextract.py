import os
import sqlite3
import sys

# Paths
cache_path = r"c:\Users\hp\OneDrive\Desktop\devine\cofsils_extraction_cache.json"
sqlite_db_path = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\mpro_reconciliation.db"
case_id = "c0f51155-e3ac-4bcf-80c2-c6603ab87716"

print("=== STARTING RESET SEQUENCE ===")

# 1. Delete the cache file to force fresh remote API calls
if os.path.exists(cache_path):
    try:
        os.remove(cache_path)
        print(f"Removed cache file: {cache_path}")
    except Exception as e:
        print(f"Failed to remove cache: {e}")
else:
    print("No cache file found, starting fresh.")

# 2. Clear local SQLite database records for this case
if os.path.exists(sqlite_db_path):
    print(f"Clearing case records in local SQLite database: {sqlite_db_path}...")
    try:
        conn = sqlite3.connect(sqlite_db_path)
        cursor = conn.cursor()
        
        tables = [
            "po_records", "agency_invoice_records", "third_party_invoice_records",
            "third_party_monitoring_records", "media_schedule_records", 
            "program_records", "pr_records", "uploaded_files"
        ]
        
        for table in tables:
            cursor.execute(f"DELETE FROM {table} WHERE case_id = ?", (case_id,))
            print(f"  Cleared {table}")
            
        cursor.execute("DELETE FROM reconciliation_cases WHERE id = ?", (case_id,))
        print("  Cleared reconciliation_cases")
        
        conn.commit()
        conn.close()
        print("SQLite Database reset completed successfully.")
    except Exception as e:
        print(f"Failed to reset SQLite database: {e}")
else:
    print(f"SQLite database not found at {sqlite_db_path}")

# 3. Trigger batch processing and sync
print("\n=== TRIGGERING BATCH EXTRACTION AND DB SYNC ===")
sys.path.append(r"c:\Users\hp\OneDrive\Desktop\devine")
import process_cofsils_batch

process_cofsils_batch.run_batch()
