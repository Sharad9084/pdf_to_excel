# 🚀 Deployment Guide for PDF Data Extractor API

## Overview
Your PDF Data Extractor has been converted into a FastAPI-based REST API that can be deployed to the cloud for free.

## Files Created

1. **api.py** - FastAPI server with 5 main endpoints
2. **requirements.txt** - Python dependencies
3. **Procfile** - Deployment configuration
4. **runtime.txt** - Python version specification
5. **.gitignore** - Git ignore rules
6. **DEPLOYMENT.md** - This file

## API Endpoints

### 1. Single PDF Extraction

```bash
# Extract Purchase Order
curl -X POST "http://localhost:8000/api/extract-po" \
  -F "file=@po.pdf"

# Extract Agency Invoice
curl -X POST "http://localhost:8000/api/extract-agency-invoice" \
  -F "file=@agency_invoice.pdf"

# Extract Broadcaster Invoice
curl -X POST "http://localhost:8000/api/extract-broadcaster-invoice" \
  -F "file=@broadcaster_invoice.pdf" \
  -F "api_key=YOUR_GEMINI_API_KEY"

# Extract Monitoring Report
curl -X POST "http://localhost:8000/api/extract-monitoring" \
  -F "file=@monitoring.pdf"
```

### 2. Batch Extract (Multiple PDFs to Excel)

```bash
curl -X POST "http://localhost:8000/api/batch-extract" \
  -F "files=@file1.pdf" \
  -F "files=@file2.pdf" \
  -F "files=@file3.pdf" \
  -F "api_key=YOUR_GEMINI_API_KEY"
```

Response:
```json
{
  "status": "success",
  "message": "Batch extraction completed",
  "summary": {
    "po_rows": 1,
    "agency_headers": 2,
    "agency_spots": 50,
    "broadcaster_rows": 10,
    "monitoring_rows": 100,
    "unknown": 0
  },
  "download_url": "/api/download/extracted_data_xyz.xlsx"
}
```

## Local Testing

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

2. Run API locally:
   ```bash
   python api.py
   ```
   Or with uvicorn directly:
   ```bash
   uvicorn api:app --reload --host 0.0.0.0 --port 8000
   ```

3. Access API:
   - Swagger UI: http://localhost:8000/docs
   - ReDoc: http://localhost:8000/redoc
   - API Docs: http://localhost:8000/api/docs

## Deployment on Render

### Step 1: Create GitHub Repository

1. Initialize git (if not already done):
   ```bash
   git init
   git add .
   git commit -m "Add FastAPI server and deployment files"
   git remote add origin https://github.com/YOUR_USERNAME/devine.git
   git push -u origin main
   ```

2. If using GitHub CLI:
   ```bash
   gh repo create devine --source=. --remote=origin --push
   ```

### Step 2: Deploy on Render

1. Go to https://render.com
2. Sign up with GitHub account
3. Click "New Web Service"
4. Select "Deploy an existing repository"
5. Search for and select "devine" repository
6. Fill in the form:
   - **Name**: pdf-extractor-api (or any name)
   - **Environment**: Python 3.11
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `uvicorn api:app --host 0.0.0.0 --port $PORT`
7. Under "Environment", add:
   - `GEMINI_API_KEY`: Your Gemini API key (optional, can be passed per request)
8. Click "Create Web Service"

### Step 3: Get Your API URL

After deployment, Render will give you a URL like:
```
https://pdf-extractor-api.onrender.com
```

## Environment Variables

Set these in your Render dashboard under "Environment":

- `GEMINI_API_KEY` (optional): Your Google Gemini API key for Broadcaster Invoice extraction
- `PORT` (auto-set by Render): The port your app runs on

## Using the Deployed API

Once deployed, replace `localhost:8000` with your Render URL:

```bash
# Example with deployed API
curl -X POST "https://pdf-extractor-api.onrender.com/api/batch-extract" \
  -F "files=@po.pdf" \
  -F "files=@invoice.pdf" \
  -F "api_key=YOUR_GEMINI_API_KEY"
```

## Alternative Hosting Options

### Heroku (with credit card, small cost)
```bash
heroku login
heroku create your-app-name
git push heroku main
heroku config:set GEMINI_API_KEY=your_key_here
```

### Railway (https://railway.app)
- Connect GitHub repo
- Railway auto-detects Procfile
- Deploy with one click

### Replit (https://replit.com)
- Import from GitHub
- Select Python as language
- Run with `python api.py`

## API Response Examples

### Success Response
```json
{
  "status": "success",
  "message": "PO extracted successfully",
  "data": {
    "file": "po_001.pdf",
    "Advertiser_Name": "Xiaomi Technology India",
    "PO_Number": "240510242B",
    "PO_Date": "2024-05-10",
    "PO_Amount_Incl_Tax": 100000.00
  }
}
```

### Error Response
```json
{
  "status": "error",
  "message": "Extraction failed: File format not recognized",
  "data": null
}
```

## Troubleshooting

### File size too large
- Render has 30GB disk space, but PDFs are usually small
- For very large batches, split into multiple requests

### Timeout on large batches
- Default timeout is ~30 seconds for Free tier
- For batch processing, use upgrade to Standard plan

### Gemini API rate limit
- Add delays between requests
- Use free tier with rate limiting: 15 requests/minute

## Costs

- **Render Free Tier**: 
  - No uptime guarantee (sleeps after inactivity)
  - Suitable for low-traffic demos
  
- **Render Paid ($12/month)**:
  - 24/7 uptime
  - Good for production use

- **Gemini API**: 
  - Free tier: 15 requests/minute
  - Paid: Based on usage

## Security Notes

1. Never commit `.env` files or API keys
2. Use environment variables for sensitive data
3. Consider adding API key authentication for production
4. Rate limit your endpoints

## Next Steps

1. Test locally with `python api.py`
2. Push to GitHub
3. Deploy on Render
4. Share the URL with your team
5. Monitor API usage in Render dashboard

## Support

For issues:
- FastAPI docs: https://fastapi.tiangolo.com
- Render docs: https://render.com/docs
- Gemini API docs: https://ai.google.dev

---

**Your API is now ready for the cloud! 🎉**
