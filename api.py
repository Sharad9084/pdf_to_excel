"""
FastAPI server for PDF Data Extractor.
Exposes HTTP endpoints to process media campaign documents.
"""

import os
import tempfile
import shutil
import io
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import pandas as pd

from extractor import (
    classify_pdf, extract_po, extract_agency_invoice,
    extract_broadcaster_invoice, extract_monitoring,
    load_templates, save_templates
)


app = FastAPI(
    title="PDF Data Extractor API",
    description="Extract structured data from media campaign PDFs",
    version="1.0.0"
)

# Enable CORS for all origins (for web frontend)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


class ExtractResponse(BaseModel):
    """Response for extraction"""
    status: str
    message: str
    data: Optional[dict] = None


class BatchResponse(BaseModel):
    """Response for batch extraction"""
    status: str
    message: str
    summary: Optional[dict] = None
    download_url: Optional[str] = None


# Global temp dir for batch processing
TEMP_DIR = tempfile.gettempdir()
UPLOAD_DIR = os.path.join(TEMP_DIR, "pdf_uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)


def cleanup_temp_file(filepath: str):
    """Background task to cleanup temp files"""
    try:
        if os.path.exists(filepath):
            os.remove(filepath)
    except Exception:
        pass


@app.get("/")
def health_check():
    """Health check endpoint"""
    return {"status": "ok", "message": "PDF Data Extractor API is running"}


@app.post("/api/extract-po", response_model=ExtractResponse)
async def extract_po_endpoint(file: UploadFile = File(...)):
    """Extract data from Purchase Order PDF"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    try:
        # Save temp file
        temp_path = os.path.join(UPLOAD_DIR, file.filename)
        with open(temp_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        # Extract data
        result = extract_po(temp_path)
        
        # Cleanup
        os.remove(temp_path)
        
        return ExtractResponse(
            status="success",
            message="PO extracted successfully",
            data=result
        )
    except Exception as e:
        return ExtractResponse(
            status="error",
            message=f"Extraction failed: {str(e)}",
            data=None
        )


@app.post("/api/extract-agency-invoice", response_model=ExtractResponse)
async def extract_agency_invoice_endpoint(file: UploadFile = File(...)):
    """Extract data from Agency Invoice PDF"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    try:
        temp_path = os.path.join(UPLOAD_DIR, file.filename)
        with open(temp_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        header, spots = extract_agency_invoice(temp_path)
        
        os.remove(temp_path)
        
        return ExtractResponse(
            status="success",
            message="Agency Invoice extracted successfully",
            data={
                "header": header,
                "spots": spots,
                "spot_count": len(spots)
            }
        )
    except Exception as e:
        return ExtractResponse(
            status="error",
            message=f"Extraction failed: {str(e)}",
            data=None
        )


@app.post("/api/extract-broadcaster-invoice", response_model=ExtractResponse)
async def extract_broadcaster_invoice_endpoint(
    file: UploadFile = File(...),
    api_key: Optional[str] = None
):
    """Extract data from Broadcaster Invoice PDF"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    try:
        temp_path = os.path.join(UPLOAD_DIR, file.filename)
        with open(temp_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        # Get API key from env if not provided
        api_key = api_key or os.environ.get("GEMINI_API_KEY")
        
        # Load templates
        templates = load_templates("templates.json")
        
        # Extract
        results = extract_broadcaster_invoice(temp_path, templates, "templates.json", api_key)
        
        os.remove(temp_path)
        
        return ExtractResponse(
            status="success",
            message="Broadcaster Invoice extracted successfully",
            data={"rows": results, "count": len(results)}
        )
    except Exception as e:
        return ExtractResponse(
            status="error",
            message=f"Extraction failed: {str(e)}",
            data=None
        )


@app.post("/api/extract-monitoring", response_model=ExtractResponse)
async def extract_monitoring_endpoint(file: UploadFile = File(...)):
    """Extract data from Monitoring Report PDF"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(status_code=400, detail="File must be a PDF")
    
    try:
        temp_path = os.path.join(UPLOAD_DIR, file.filename)
        with open(temp_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)
        
        rows = extract_monitoring(temp_path)
        
        os.remove(temp_path)
        
        return ExtractResponse(
            status="success",
            message="Monitoring Report extracted successfully",
            data={"rows": rows, "count": len(rows)}
        )
    except Exception as e:
        return ExtractResponse(
            status="error",
            message=f"Extraction failed: {str(e)}",
            data=None
        )


@app.post("/api/batch-extract", response_model=BatchResponse)
async def batch_extract(
    files: list[UploadFile] = File(...),
    api_key: Optional[str] = None,
    background_tasks: BackgroundTasks = BackgroundTasks()
):
    """
    Batch extract from multiple PDFs and return Excel file
    Process all PDFs and combine results into one Excel
    """
    if not files:
        raise HTTPException(status_code=400, detail="No files provided")
    
    # Validate all files are PDFs
    for file in files:
        if not file.filename.lower().endswith('.pdf'):
            raise HTTPException(status_code=400, detail=f"File {file.filename} must be a PDF")
    
    try:
        # Create a temp working directory
        work_dir = os.path.join(UPLOAD_DIR, f"batch_{os.urandom(8).hex()}")
        os.makedirs(work_dir, exist_ok=True)
        
        # Save all files
        for file in files:
            filepath = os.path.join(work_dir, file.filename)
            with open(filepath, "wb") as buffer:
                content = await file.read()
                buffer.write(content)
        
        # Process all PDFs
        api_key = api_key or os.environ.get("GEMINI_API_KEY")
        templates = load_templates("templates.json")
        
        pos = []
        ag_headers = []
        ag_spots = []
        br_invoices = []
        mon_rows = []
        unknown = []
        
        for filename in os.listdir(work_dir):
            filepath = os.path.join(work_dir, filename)
            if not os.path.isfile(filepath):
                continue
            
            try:
                doc_type = classify_pdf(filepath)
                
                if doc_type == 'po':
                    pos.append(extract_po(filepath))
                elif doc_type == 'agency_invoice':
                    h, sp = extract_agency_invoice(filepath)
                    ag_headers.append(h)
                    ag_spots.extend(sp)
                elif doc_type == 'broadcaster_invoice':
                    br_invoices.extend(extract_broadcaster_invoice(filepath, templates, "templates.json", api_key))
                elif doc_type == 'monitoring':
                    mon_rows.extend(extract_monitoring(filepath))
                else:
                    unknown.append({'file': filename})
            except Exception as e:
                unknown.append({'file': filename, 'error': str(e)})
        
        # Create Excel output
        output_file = os.path.join(work_dir, "extracted_data.xlsx")
        with pd.ExcelWriter(output_file, engine='openpyxl') as xl:
            if pos:
                pd.DataFrame(pos).to_excel(xl, sheet_name='1_PO', index=False)
            if ag_headers:
                pd.DataFrame(ag_headers).to_excel(xl, sheet_name='2_Agency_Invoice', index=False)
            if ag_spots:
                pd.DataFrame(ag_spots).to_excel(xl, sheet_name='2_Agency_Spots', index=False)
            if br_invoices:
                pd.DataFrame(br_invoices).to_excel(xl, sheet_name='3_Broadcaster_Invoice', index=False)
            if mon_rows:
                pd.DataFrame(mon_rows).to_excel(xl, sheet_name='4_Monitoring', index=False)
            if unknown:
                pd.DataFrame(unknown).to_excel(xl, sheet_name='Unknown', index=False)
        
        # Schedule cleanup
        background_tasks.add_task(shutil.rmtree, work_dir, ignore_errors=True)
        
        summary = {
            "po_rows": len(pos),
            "agency_headers": len(ag_headers),
            "agency_spots": len(ag_spots),
            "broadcaster_rows": len(br_invoices),
            "monitoring_rows": len(mon_rows),
            "unknown": len(unknown)
        }
        
        return BatchResponse(
            status="success",
            message="Batch extraction completed",
            summary=summary,
            download_url=f"/api/download/{os.path.basename(output_file)}"
        )
    except Exception as e:
        return BatchResponse(
            status="error",
            message=f"Batch extraction failed: {str(e)}",
            summary=None,
            download_url=None
        )


@app.get("/api/download/{filename}")
async def download_file(filename: str, background_tasks: BackgroundTasks):
    """Download extracted data file"""
    filepath = os.path.join(UPLOAD_DIR, filename)
    if not os.path.exists(filepath):
        raise HTTPException(status_code=404, detail="File not found")
    
    # Schedule cleanup after sending
    background_tasks.add_task(cleanup_temp_file, filepath)
    
    return FileResponse(
        path=filepath,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=filename
    )


@app.get("/api/docs")
def api_docs():
    """API documentation"""
    return {
        "title": "PDF Data Extractor API",
        "version": "1.0.0",
        "endpoints": {
            "POST /api/extract-po": "Extract Purchase Order data",
            "POST /api/extract-agency-invoice": "Extract Agency Invoice data",
            "POST /api/extract-broadcaster-invoice": "Extract Broadcaster Invoice data",
            "POST /api/extract-monitoring": "Extract Monitoring Report data",
            "POST /api/batch-extract": "Batch extract from multiple PDFs into Excel"
        }
    }


if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8000))
    uvicorn.run(app, host="0.0.0.0", port=port)
