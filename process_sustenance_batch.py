import os
import sys
import uuid
import datetime
import random
import json

# Add current workspace and backend directory to path
sys.path.append(r"c:\Users\hp\OneDrive\Desktop\devine")
sys.path.append(r"C:\Users\hp\OneDrive\Desktop\fronted\backend")

import extractor
import server

# Source files directory
sustenance_path = r"c:\Users\hp\OneDrive\Desktop\devine\Sustenance_Redmi_Note_13_May24\Sustenance_Redmi_Note_13_May24"
templates_path = r"c:\Users\hp\OneDrive\Desktop\devine\templates.json"

def run_batch_processing():
    print("=== STARTING BATCH PDF EXTRACTION & INTEGRATION ===")
    print(f"Folder: {sustenance_path}")
    
    if not os.path.exists(sustenance_path):
        print("[ERROR] Sustenance folder does not exist.")
        return
        
    # Load templates
    templates = extractor.load_templates(templates_path)
    
    # Walk through folders and find all PDFs
    pdf_files = []
    for root, dirs, files in os.walk(sustenance_path):
        for f in files:
            if f.lower().endswith('.pdf'):
                pdf_files.append(os.path.join(root, f))
    
    pdf_files.sort()
    total_files = len(pdf_files)
    print(f"Found {total_files} PDF files to process.\n")
    
    # Target dataset mapping
    datasets = {
        "po": [],
        "mediaSchedule": [],
        "agency": [],
        "thirdPartyInvoice": [],
        "thirdPartyMonitoring": []
    }
    
    # Metadata config
    metadata = {
        "agency": "GroupM Media India Private Limited",
        "advertiser": "Xiaomi",
        "medium": "TV",
        "campaignStartDate": "2024-05-01",
        "campaignPeriod": "May 2024"
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
    
    # Loop over files and extract data
    for idx, path in enumerate(pdf_files, 1):
        filename = os.path.basename(path)
        rel_path = os.path.relpath(path, sustenance_path)
        
        # Classify document
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
            print(f"[{idx}/{total_files}] [SKIP] {rel_path} - Unknown layout type")
            failed_count += 1
            continue
            
        print(f"[{idx}/{total_files}] [PROCESS] {rel_path} -> {source_key}...", end=" ", flush=True)
        
        try:
            rows = []
            if source_key == "po":
                row_data = extractor.extract_po(path)
                rows.append(row_data)
            elif source_key == "agency":
                header, spots = extractor.extract_agency_invoice(path)
                for spot in spots:
                    # Merge header and spot details
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
                
            datasets[source_key].extend(enriched)
            success_count += 1
            print(f"[OK] Extracted {len(enriched)} rows")
            
        except Exception as e:
            failed_count += 1
            print(f"[ERROR] {e}")
            
    print(f"\nExtraction complete: {success_count} success, {failed_count} failed")
    
    # Save the case to the SQLite database
    print("\nIntegrating extracted rows into database...")
    case_id = str(uuid.uuid4())
    case_name = "Sustenance Redmi Note 13 May 2024"
    
    case_payload = {
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
    
    try:
        # Initialize db schemas if not present
        server.init_db()
        
        # Save case
        server.upsert_case(case_payload)
        print(f"[DB OK] Saved case: '{case_name}' with ID: {case_id}")
        print(f"Total Rows Saved:")
        print(f"  - PO: {len(datasets['po'])}")
        print(f"  - Agency Invoice: {len(datasets['agency'])}")
        print(f"  - Broadcaster Invoice: {len(datasets['thirdPartyInvoice'])}")
        print(f"  - Monitoring: {len(datasets['thirdPartyMonitoring'])}")
        
        # Start server.py in background if not running so frontend can fetch it
        print("\nAll tasks completed successfully!")
    except Exception as e:
        print(f"[DB ERROR] Failed to save case: {e}")

if __name__ == "__main__":
    run_batch_processing()
