# PDF Data Extractor — POC

Free, offline extractor for media-campaign documents (PO, Agency Invoice,
Broadcaster Invoice, Monitoring Report).

## How to run

```bash
pip install pymupdf pdfplumber openpyxl pandas

python3 extractor.py \
    --input  /path/to/Sustenance_Redmi_Note_13_May24 \
    --output extracted_data.xlsx
```

The script walks the folder, auto-classifies each PDF, and writes one Excel
file with these sheets:

| Sheet | Rows in this run |
|---|---|
| 1_PO | 1 |
| 2_Agency_Invoice (header) | 30 |
| 2_Agency_Spots (spot-level lines) | 1,227 |
| 3_Broadcaster_Invoice | 240 |
| 4_Monitoring | 11,762 |

## Fill quality on the sample ZIP

| Sheet | Fill rate |
|---|---|
| PO | 100% |
| Agency Invoice header | 100% |
| Agency Spots | 100% |
| Broadcaster Invoice | 68.6% |
| Monitoring | 100% |

The lower broadcaster fill rate is expected: the sample includes 14+ different
broadcaster layouts (Polimer, Writemen / Public TV, Matrix, Sun, ZEE, ABP, Star,
TV18, Network18, etc.). The current generic parser covers the most common
layouts; SAP-style multi-column layouts (e.g. Polimer) need per-broadcaster
templates or an LLM fallback.

## Recommended next steps

1. **Per-broadcaster templates** — one small parser per broadcaster goes into a
   registry keyed by GSTIN / PAN. Once a broadcaster has a template, parsing is
   instant, deterministic, and free.
2. **LLM fallback (free tier)** for any new broadcaster:
   Google Gemini 2.0 Flash free tier or Groq Llama 3.3 free tier — pass the
   raw page text + a JSON schema, get extracted fields back.
3. **Human-review UI** (Streamlit) — operator reviews/corrects LLM output, then
   the corrected fields are auto-saved as a new template. Next month's invoice
   from the same broadcaster runs through the template for free.

## File layout

```
pdf_extractor/
  extractor.py         # main script (all logic in one file for POC)
  README.md            # this file
  extracted_data.xlsx  # output
```
