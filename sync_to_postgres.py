import sqlite3
import json
import urllib.request
import urllib.parse
import time
import sys

# Paths
sqlite_db_path = r"C:\Users\hp\OneDrive\Desktop\fronted\backend\mpro_reconciliation.db"
vercel_url = "https://mpro-ai.vercel.app"
case_id = "e3ac65e8-3ab1-4bcf-80c2-c6603ab87716"

print("=== SYNCING CASE TO CLOUD POSTGRES IN CHUNKS ===")

# 1. Fetch case from local SQLite
print(f"Reading case '{case_id}' from SQLite database...")
try:
    conn = sqlite3.connect(sqlite_db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    cursor.execute("SELECT raw_json FROM reconciliation_cases WHERE id = ?", (case_id,))
    row = cursor.fetchone()
    if not row:
        print(f"Error: Case '{case_id}' not found in local SQLite database.")
        sys.exit(1)
    case_data = json.loads(row["raw_json"])
    conn.close()
    print("  Successfully read case from SQLite.")
except Exception as e:
    print(f"  Error reading local SQLite: {e}")
    sys.exit(1)

# Separate datasets from metadata
datasets = case_data.pop("datasets", {})
print(f"Campaign dataset sizes:")
for k, v in datasets.items():
    print(f"  - {k}: {len(v or [])} rows")

# 2. Authenticate on Vercel
login_url = vercel_url + "/api/auth/signin"
login_payload = json.dumps({
    "username": "auditor@mpro.com",
    "password": "Auditor@2026"
}).encode('utf-8')

print(f"\nSigning in to Vercel: {login_url}...")
token = None
try:
    req = urllib.request.Request(
        login_url, 
        data=login_payload, 
        headers={'Content-Type': 'application/json'}
    )
    with urllib.request.urlopen(req, timeout=15) as response:
        res = json.loads(response.read().decode('utf-8'))
        token = res["token"]
        print(f"  Sign in successful! Token: {token[:10]}...")
except Exception as e:
    print(f"  Sign in failed: {e}")
    sys.exit(1)

# helper function for POST requests with auth token
def api_post(endpoint, payload):
    url = vercel_url + endpoint
    body = json.dumps(payload, ensure_ascii=False).encode('utf-8')
    req = urllib.request.Request(
        url,
        data=body,
        headers={
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {token}'
        }
    )
    with urllib.request.urlopen(req, timeout=45) as response:
        return json.loads(response.read().decode('utf-8'))

# 3. Initialize case on Vercel
print("\n[Step 1/3] Initializing campaign case on Vercel...")
try:
    init_res = api_post("/api/cases/init", {"case": case_data})
    print(f"  Initialization OK. Case ID: {init_res.get('id')}")
except Exception as e:
    print(f"  Failed to initialize case: {e}")
    sys.exit(1)

# 4. Upload datasets in chunks
print("\n[Step 2/3] Uploading dataset rows in chunks of 2,000...")
CHUNK_SIZE = 2000

for source_key, rows in datasets.items():
    if not rows:
        continue
    total_rows = len(rows)
    print(f"  Processing '{source_key}' ({total_rows} rows)...")
    
    for i in range(0, total_rows, CHUNK_SIZE):
        chunk = rows[i:i + CHUNK_SIZE]
        print(f"    Uploading rows {i+1} to {min(i + CHUNK_SIZE, total_rows)}...", end=" ", flush=True)
        start_t = time.time()
        try:
            api_post("/api/cases/chunk", {
                "case_id": case_id,
                "source": source_key,
                "rows": chunk
            })
            print(f"OK ({time.time() - start_t:.1f}s)")
        except Exception as e:
            print(f"\n    Failed to upload chunk: {e}")
            sys.exit(1)

# 5. Finalize the case
print("\n[Step 3/3] Finalizing case on Vercel PostgreSQL...")
try:
    start_t = time.time()
    finalize_res = api_post("/api/cases/finalize", {"case_id": case_id})
    print(f"  Finalization successful in {time.time() - start_t:.1f}s! Status: {finalize_res.get('status')}")
except Exception as e:
    print(f"  Failed to finalize case: {e}")
    sys.exit(1)

print("\nSync completed successfully! The entire campaign dataset is now fully synced to Vercel PostgreSQL.")
