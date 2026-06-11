import os
import sys
import uuid
import datetime
import random
import json
import urllib.request
import urllib.parse
import time
import sqlite3

# Add path for extractor and backend
sys.path.append(r"c:\Users\hp\OneDrive\Desktop\devine")
sys.path.append(r"C:\Users\hp\OneDrive\Desktop\fronted\backend")

import extractor

# Directories
zip_dir = r"c:\Users\hp\OneDrive\Desktop\devine\cofsils_data"
rar_dir = r"c:\Users\hp\OneDrive\Desktop\devine\cofsils_extracted_rar"
templates_path = r"c:\Users\hp\OneDrive\Desktop\devine\templates.json"

case_id = "c0f51155-e3ac-4bcf-80c2-c6603ab87716"
case_name = "Cipla Cofsils Campaign"

def get_unique_pdfs():
    unique_pdfs = {}
    
    # Walk zip dir
    if os.path.exists(zip_dir):
        for root, dirs, files in os.walk(zip_dir):
            for f in files:
                if f.lower().endswith('.pdf'):
                    p = os.path.join(root, f)
                    rel = os.path.relpath(p, zip_dir)
                    rel_parts = rel.split(os.sep)
                    if rel_parts[0] in ["3. Cofsils", "Cofsils Broadcaster Invoices and Tc"]:
                        rel_norm = os.path.join(*rel_parts[1:])
                    else:
                        rel_norm = rel
                    unique_pdfs[rel_norm] = p
                    
    # Walk rar dir
    if os.path.exists(rar_dir):
        for root, dirs, files in os.walk(rar_dir):
            for f in files:
                if f.lower().endswith('.pdf'):
                    p = os.path.join(root, f)
                    rel = os.path.relpath(p, rar_dir)
                    rel_parts = rel.split(os.sep)
                    if rel_parts[0] in ["3. Cofsils", "Cofsils Broadcaster Invoices and Tc"]:
                        rel_norm = os.path.join(*rel_parts[1:])
                    else:
                        rel_norm = rel
                    if rel_norm not in unique_pdfs:
                        unique_pdfs[rel_norm] = p
                        
    return unique_pdfs

def run_batch():
    print("=== STARTING CIPLA COFSILS BATCH EXTRACTION ===")
    
    unique_pdfs = get_unique_pdfs()
    total_files = len(unique_pdfs)
    print(f"Found {total_files} unique PDF files to process.")
    
    templates = extractor.load_templates(templates_path)
    
    # Load cache
    cache_path = r"c:\Users\hp\OneDrive\Desktop\devine\cofsils_extraction_cache.json"
    cache = {}
    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'r', encoding='utf-8') as f:
                cache = json.load(f)
            print(f"Loaded {len(cache)} files from cache.")
        except Exception as e:
            print(f"Failed to load cache: {e}")

    # Target dataset mapping
    datasets = {
        "po": [],
        "mediaSchedule": [],
        "agency": [],
        "thirdPartyInvoice": [],
        "thirdPartyMonitoring": []
    }
    
    metadata = {
        "agency": "Madison Communications Pvt Ltd",
        "advertiser": "Cipla Health Ltd",
        "medium": "TV",
        "campaignStartDate": "2024-11-01",
        "campaignPeriod": "Nov 2024 - Dec 2025"
    }
    
    source_prefixes = {
        "po": "po",
        "mediaSchedule": "sch",
        "agency": "agi",
        "thirdPartyInvoice": "bri",
        "thirdPartyMonitoring": "moi"
    }
    
    source_labels = {
        "po": "Purchase Order",
        "mediaSchedule": "Media Estimate/Schedule",
        "agency": "Agency Invoice",
        "thirdPartyInvoice": "Publisher Invoice",
        "thirdPartyMonitoring": "3rd Party Monitoring"
    }
    
    success_count = 0
    failed_count = 0
    skipped_count = 0
    
    # Sort the items for deterministic order
    sorted_items = sorted(unique_pdfs.items())
    
    for idx, (rel_norm, path) in enumerate(sorted_items, 1):
        filename = os.path.basename(path)
        
        # Sniff layout
        kind = extractor.classify_pdf(path)
        
        source_key = None
        if kind == "po":
            source_key = "po"
        elif kind == "agency_invoice":
            source_key = "agency"
        elif kind == "broadcaster_invoice":
            source_key = "thirdPartyInvoice"
        elif kind == "monitoring":
            source_key = "thirdPartyMonitoring"
            
        if not source_key:
            print(f"[{idx}/{total_files}] [SKIP] {rel_norm} - Unknown layout type")
            skipped_count += 1
            continue
            
        # Check cache
        file_mtime = str(os.path.getmtime(path))
        cache_key = f"{rel_norm}:{file_mtime}"
        
        if cache_key in cache:
            rows = cache[cache_key]
            # Apply metadata enrichment and add to datasets
            datasets[source_key].extend(rows)
            success_count += 1
            print(f"[{idx}/{total_files}] [CACHED] {rel_norm} -> {source_key}... [OK] Loaded {len(rows)} rows")
            continue
            
        print(f"[{idx}/{total_files}] [PROCESS] {rel_norm} -> {source_key}...", end=" ", flush=True)
        
        start_time = time.time()
        try:
            rows = []
            if source_key == "po":
                row_data = extractor.extract_po(path)
                rows.append(row_data)
            elif source_key == "agency":
                header, spots = extractor.extract_agency_invoice(path)
                for spot in spots:
                    rows.append({**header, **spot})
            elif source_key == "thirdPartyInvoice":
                spots = extractor.extract_broadcaster_invoice(path, templates, templates_path)
                rows.extend(spots)
            elif source_key == "thirdPartyMonitoring":
                mon_rows = extractor.extract_monitoring(path)
                rows.extend(mon_rows)
                
            # Apply metadata enrichment
            prefix = source_prefixes[source_key]
            stamp = f"{datetime.datetime.now().strftime('%y%m%d%H%M')}-{random.randint(1000, 9999)}"
            
            enriched = []
            for r_idx, r in enumerate(rows, 1):
                next_row = {**r}
                next_row["Import ID"] = next_row.get("Import ID") or f"{prefix}-{stamp}-{str(r_idx).zfill(3)}"
                next_row["Document Type"] = next_row.get("Document Type") or source_labels[source_key]
                next_row["Agency Name"] = next_row.get("Agency Name") or metadata["agency"]
                next_row["Advertiser Name"] = next_row.get("Advertiser Name") or metadata["advertiser"]
                next_row["Medium"] = next_row.get("Medium") or metadata["medium"]
                next_row["Media Type"] = next_row.get("Media Type") or metadata["medium"]
                next_row["Campaign Start Date"] = next_row.get("Campaign Start Date") or metadata["campaignStartDate"]
                next_row["Campaign End Date"] = ""
                next_row["Campaign Period"] = next_row.get("Campaign Period") or metadata["campaignPeriod"]
                next_row["File Name"] = filename
                enriched.append(next_row)
                
            # Save to cache
            cache[cache_key] = enriched
            if idx % 10 == 0 or idx == total_files:
                try:
                    with open(cache_path, 'w', encoding='utf-8') as f:
                        json.dump(cache, f, ensure_ascii=False, indent=2)
                except Exception:
                    pass
                    
            datasets[source_key].extend(enriched)
            success_count += 1
            print(f"[OK] Extracted {len(enriched)} rows ({time.time() - start_time:.1f}s)")
            
        except Exception as e:
            failed_count += 1
            print(f"[ERROR] {e} ({time.time() - start_time:.1f}s)")
            
    print(f"\nBatch processing finished: {success_count} success, {failed_count} failed, {skipped_count} skipped")
    print(f"Total Rows Extracted:")
    print(f"  - PO: {len(datasets['po'])} rows")
    print(f"  - Agency Invoice: {len(datasets['agency'])} rows")
    print(f"  - Broadcaster Invoice: {len(datasets['thirdPartyInvoice'])} rows")
    print(f"  - Monitoring: {len(datasets['thirdPartyMonitoring'])} rows")
    
    # ----------------------------------------------------
    # Sync to Vercel PostgreSQL
    # ----------------------------------------------------
    print("\n=== SYNCING TO CLOUD POSTGRES ===")
    vercel_url = "https://mpro-ai.vercel.app"
    
    login_url = vercel_url + "/api/auth/signin"
    login_payload = json.dumps({
        "username": "auditor@mpro.com",
        "password": "Auditor@2026"
    }).encode('utf-8')
    
    print(f"Signing in to Vercel: {login_url}...")
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
        return
        
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
            
    case_payload = {
        "id": case_id,
        "name": case_name,
        "activeView": "reconciliation",
        "columnOrders": {},
        "columnWidths": {},
        "sort": {}
    }
    
    print("\n[Step 1/3] Initializing campaign case on Vercel...")
    try:
        init_res = api_post("/api/cases/init", {"case": case_payload})
        print(f"  Initialization OK. Case ID: {init_res.get('id')}")
    except Exception as e:
        print(f"  Failed to initialize case on Vercel: {e}")
        return
        
    print("\n[Step 2/3] Uploading dataset rows in chunks of 2,000...")
    CHUNK_SIZE = 500
    for source_key, rows in datasets.items():
        if not rows:
            continue
        total_rows = len(rows)
        print(f"  Uploading '{source_key}' ({total_rows} rows)...")
        for i in range(0, total_rows, CHUNK_SIZE):
            chunk = rows[i:i + CHUNK_SIZE]
            print(f"    Chunk {i//CHUNK_SIZE + 1} ({i+1} to {min(i + CHUNK_SIZE, total_rows)})...", end=" ", flush=True)
            start_t = time.time()
            try:
                api_post("/api/cases/chunk", {
                    "case_id": case_id,
                    "source": source_key,
                    "rows": chunk
                })
                print(f"OK ({time.time() - start_t:.1f}s)")
                time.sleep(0.5)
            except Exception as e:
                print(f"\n    Failed to upload chunk: {e}")
                return
                
    print("\n[Step 3/3] Finalizing case on Vercel PostgreSQL...")
    try:
        start_t = time.time()
        finalize_res = api_post("/api/cases/finalize", {"case_id": case_id})
        print(f"  Finalization successful in {time.time() - start_t:.1f}s! Status: {finalize_res.get('status')}")
    except Exception as e:
        print(f"  Failed to finalize case on Vercel: {e}")
        return
        
    # ----------------------------------------------------
    # Also save to local SQLite database for offline consistency
    # ----------------------------------------------------
    print("\n=== SAVING TO LOCAL SQLITE ===")
    try:
        # Import server to use upsert_case
        sys.path.append(r"C:\Users\hp\OneDrive\Desktop\fronted\backend")
        import server
        
        local_payload = {
            "case": {
                "id": case_id,
                "name": case_name,
                "datasets": datasets,
                "activeView": "reconciliation",
                "columnOrders": {},
                "columnWidths": {},
                "sort": {}
            },
            "actor": "auditor@mpro.com"
        }
        server.init_db()
        server.upsert_case(local_payload)
        print("  Local SQLite save OK.")
    except Exception as e:
        print(f"  Failed to save locally: {e}")
        
    print("\nAll batch processing and sync tasks finished successfully!")

if __name__ == "__main__":
    run_batch()
