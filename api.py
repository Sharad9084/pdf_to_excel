from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
import os
import shutil
import tempfile
import sys

# Add current workspace directory to system path to import extractor.py
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# Import our extraction functions and templates registry helpers
import extractor

app = FastAPI(
    title="Media Campaign PDF Data Extractor API",
    description="Offline microservice to parse structured details and spot-logs from media campaign PDFs.",
    version="1.0.0"
)

# Enable CORS so any external software/front-end can connect
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

TEMPLATES_PATH = "templates.json"
templates = extractor.load_templates(TEMPLATES_PATH)

@app.get("/")
def home():
    return {
        "message": "Welcome to the Media Campaign PDF Data Extractor API",
        "endpoints": {
            "swagger_docs": "/docs",
            "redoc": "/redoc",
            "extract": "POST /extract"
        }
    }

@app.post("/extract")
async def extract_pdf(file: UploadFile = File(...)):
    # Verify file is a PDF
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="Only PDF files are supported.")
        
    # Write to a temporary file
    temp_dir = tempfile.gettempdir()
    temp_file_path = os.path.join(temp_dir, f"temp_{file.filename}")
    
    try:
        with open(temp_file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
            
        # Classify the PDF
        kind = extractor.classify_pdf(temp_file_path)
        
        # Parse depending on document type
        if kind == 'po':
            data = extractor.extract_po(temp_file_path)
            return {
                "status": "success",
                "document_type": "po",
                "file_name": file.filename,
                "data": data
            }
            
        elif kind == 'agency_invoice':
            header, spots = extractor.extract_agency_invoice(temp_file_path)
            return {
                "status": "success",
                "document_type": "agency_invoice",
                "file_name": file.filename,
                "header": header,
                "spots": spots
            }
            
        elif kind == 'broadcaster_invoice':
            # Extract spots (runs completely offline)
            spots = extractor.extract_broadcaster_invoice(temp_file_path, templates, TEMPLATES_PATH)
            
            # Simple header extraction for response overview
            header = {}
            if spots:
                # Use first spot's metadata as header
                s0 = spots[0]
                header = {
                    "Broadcaster_Name": s0.get("Broadcaster_Name"),
                    "Advertiser_Name": s0.get("Advertiser_Name"),
                    "Agency_Name": s0.get("Agency_Name"),
                    "Invoice_Number": s0.get("Invoice_Number"),
                    "Invoice_Date": s0.get("Invoice_Date"),
                    "Total_Amount": s0.get("Total_Amount"),
                    "Taxable_Amount": s0.get("Taxable_Amount"),
                    "GSTIN": s0.get("GSTIN"),
                    "PAN": s0.get("PAN"),
                    "State": s0.get("State")
                }
                
            return {
                "status": "success",
                "document_type": "broadcaster_invoice",
                "file_name": file.filename,
                "header": header,
                "spots": spots
            }
            
        elif kind == 'monitoring':
            rows = extractor.extract_monitoring(temp_file_path)
            return {
                "status": "success",
                "document_type": "monitoring",
                "file_name": file.filename,
                "rows": rows
            }
            
        else:
            return {
                "status": "unknown",
                "document_type": "unknown",
                "file_name": file.filename,
                "detail": "Document layout could not be classified. No tables matched."
            }
            
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Extraction failed: {str(e)}")
        
    finally:
        # Cleanup temp file
        if os.path.exists(temp_file_path):
            os.remove(temp_file_path)

if __name__ == '__main__':
    import uvicorn
    # Local run configurations
    uvicorn.run(app, host="0.0.0.0", port=8000)
