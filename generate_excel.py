import os
import sys
import datetime
import random
import pandas as pd

# Add current workspace and backend directory to path
sys.path.append(r"c:\Users\hp\OneDrive\Desktop\devine")
sys.path.append(r"C:\Users\hp\OneDrive\Desktop\fronted\backend")

import extractor

# Source files directory
sustenance_path = r"c:\Users\hp\OneDrive\Desktop\devine\Sustenance_Redmi_Note_13_May24\Sustenance_Redmi_Note_13_May24"
templates_path = r"c:\Users\hp\OneDrive\Desktop\devine\templates.json"
output_excel = r"c:\Users\hp\OneDrive\Desktop\devine\extracted_data.xlsx"

def run_extraction_to_excel():
    api_key = os.environ.get("GEMINI_API_KEY")
    print("=== STARTING BATCH PDF EXTRACTION TO EXCEL ===")
    print(f"Folder: {sustenance_path}")
    print(f"Output: {output_excel}")
    
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
    
    # Target dataset lists
    po_list = []
    agency_invoice_list = []
    agency_spots_list = []
    broadcaster_invoice_list = []
    monitoring_list = []
    
    success_count = 0
    failed_count = 0
    
    # Loop over files and extract data
    for idx, path in enumerate(pdf_files, 1):
        filename = os.path.basename(path)
        rel_path = os.path.relpath(path, sustenance_path)
        
        # Classify document
        kind = extractor.classify_pdf(path)
        
        if not kind:
            print(f"[{idx}/{total_files}] [SKIP] {rel_path} - Unknown layout type")
            failed_count += 1
            continue
            
        print(f"[{idx}/{total_files}] [PROCESS] {rel_path} -> {kind}...", end=" ", flush=True)
        
        try:
            if kind == "po":
                row_data = extractor.extract_po(path)
                row_data["file"] = filename
                po_list.append(row_data)
            elif kind == "agency_invoice":
                header, spots = extractor.extract_agency_invoice(path)
                header["file"] = filename
                agency_invoice_list.append(header)
                for s in spots:
                    s["file"] = filename
                    agency_spots_list.append(s)
            elif kind == "broadcaster_invoice":
                spots = extractor.extract_broadcaster_invoice(path, templates, templates_path, api_key=api_key)
                for s in spots:
                    s["file"] = filename
                    # Ensure folder name or rel path is saved as folder just like original
                    s["folder"] = os.path.basename(os.path.dirname(path))
                    broadcaster_invoice_list.append(s)
            elif kind == "monitoring":
                mon_rows = extractor.extract_monitoring(path)
                for s in mon_rows:
                    s["file"] = filename
                    monitoring_list.append(s)
                
            success_count += 1
            print("[OK]")
            
        except Exception as e:
            failed_count += 1
            print(f"[ERROR] {e}")
            
    print(f"\nExtraction complete: {success_count} success, {failed_count} failed")
    
    print("\nWriting data to Excel sheet...")
    
    try:
        # Convert to DataFrames
        df_po = pd.DataFrame(po_list)
        df_agency_invoice = pd.DataFrame(agency_invoice_list)
        df_agency_spots = pd.DataFrame(agency_spots_list)
        df_broadcaster_invoice = pd.DataFrame(broadcaster_invoice_list)
        df_monitoring = pd.DataFrame(monitoring_list)
        
        # Write to multi-sheet excel file
        with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
            if not df_po.empty:
                df_po.to_excel(writer, sheet_name='1_PO', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='1_PO', index=False)
                
            if not df_agency_invoice.empty:
                df_agency_invoice.to_excel(writer, sheet_name='2_Agency_Invoice', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='2_Agency_Invoice', index=False)
                
            if not df_agency_spots.empty:
                df_agency_spots.to_excel(writer, sheet_name='2_Agency_Spots', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='2_Agency_Spots', index=False)
                
            if not df_broadcaster_invoice.empty:
                df_broadcaster_invoice.to_excel(writer, sheet_name='3_Broadcaster_Invoice', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='3_Broadcaster_Invoice', index=False)
                
            if not df_monitoring.empty:
                df_monitoring.to_excel(writer, sheet_name='4_Monitoring', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='4_Monitoring', index=False)
                
        print(f"\n[EXCEL OK] Successfully generated: {output_excel}")
        print(f"Sheet details:")
        print(f"  - 1_PO: {df_po.shape}")
        print(f"  - 2_Agency_Invoice: {df_agency_invoice.shape}")
        print(f"  - 2_Agency_Spots: {df_agency_spots.shape}")
        print(f"  - 3_Broadcaster_Invoice: {df_broadcaster_invoice.shape}")
        print(f"  - 4_Monitoring: {df_monitoring.shape}")
        
    except Exception as e:
        print(f"[EXCEL ERROR] Failed to save Excel sheet: {e}")

if __name__ == "__main__":
    run_extraction_to_excel()
