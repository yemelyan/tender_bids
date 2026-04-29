# EIS Procurement Scraper (Program #1)

This project downloads public Latvian EIS procurement documents and builds an inventory.
It reads procurement links from a CSV, collects document file links from the EIS site,
downloads the files, and exports a CSV/XLSX inventory with clickable links.

## Architecture (simple map)

Package: `src/eis_scraper/`

- `cli.py`  
  Entry point. Orchestrates the end-to-end run and writes the inventory.

- `ingest.py`  
  Reads the CSV input and normalizes procurement URLs.

- `client.py`  
  `requests.Session()` wrapper, CSRF token extraction, GET/POST helpers,
  retry + backoff + polite throttling.

- `parse_procurement.py`  
  Extracts document rows from the procurement page:
  - from `onclick="viewDocument(...)"` when available
  - from embedded JS array `ActualDocuments_items` when rows are not rendered

- `parse_viewdocument.py`  
  Parses the ViewDocument modal HTML:
  - direct `DownloadDocumentFile` links if present
  - or embedded JS array `ViewDocumentModel_Files_items` + parameters from
    hidden input `ViewDocumentModel_JsonParams`

- `download.py`  
  Streams downloads to disk, picks filenames from headers when available,
  and retries 429/5xx with backoff. Also respects the global rate limit.

- `export.py`  
  Writes `outputs/inventory.csv` and `outputs/inventory.xlsx` with hyperlinks.

- `utils.py`  
  Rate limiting, logging, error recording (`logs/errors.jsonl`).

## Website Structure (what we parse)

EIS procurement pages do not always render document rows as normal HTML.
Important structures found in EIS pages:

1) **Procurement page**
   - A hidden input with the procurement ID:  
     `ProcurementIdentifier`  
   - A JS array with document metadata:  
     `ActualDocuments_items = [ ... ]`  
     Each item includes `DocumentTitle.Id` and `DocumentTitle.DocumentLinkTypeCode`.

2) **ViewDocument modal (returned HTML)**
   - Hidden input with parameters:  
     `ViewDocumentModel_JsonParams` (JSON)
   - A JS array listing the files:  
     `ViewDocumentModel_Files_items = [ ... ]`  
     Each item includes a file `Id` and `DocumentId`.

Using these, we build:
```
https://www.eis.gov.lv/EKEIS/Document/DownloadDocumentFile?
  Id=<DocumentId>&FileId=<FileId>&DocumentLinkTypeCode=<Code>&ProcurementIdentifier=<Id>
```

## How the scraper works (step-by-step)

1) Read procurement links from `eis_e_iepirkumi_izsludinatie_2026.csv`.
2) For each procurement:
   - GET the procurement page.
   - Extract CSRF token and document rows (from HTML or JS arrays).
3) For each document row:
   - POST to `/EKEIS/Document/ViewDocument`.
   - Extract file links from the modal HTML or embedded file arrays.
4) Download each file to:
   - `downloads/<procurement_id>/<filename>`
5) Write inventory:
   - `outputs/inventory.csv`
   - `outputs/inventory.xlsx` (with clickable file + URL links)

## Politeness and stability

- Every network action is throttled to ~1–2 requests/second with jitter.
- 429 and 5xx responses are retried with exponential backoff.
- Failures for individual documents are logged to `logs/errors.jsonl` and the
  run continues.

## Output location (current default)

Base path: `C:\Users\jazep\Documents\iepirkumi dati V1`

- Downloads: `...\downloads`
- Inventory: `...\outputs`
- Logs: `...\logs`

You can override these with CLI flags if needed.

## Usage

Install dependencies (once):
```
python -m pip install -e .
```

Run all procurements:
```
python -m eis_scraper.cli --input eis_e_iepirkumi_izsludinatie_2026.csv
```

Run a small test:
```
python -m eis_scraper.cli --limit 2
```

## Bidder Extraction Script (for inspection)

The file `extract_bidders.py` is a separate analysis utility that reads already
downloaded tender documents and extracts bidder-level opening protocol data.

### Intention

- Provide a reproducible way to identify bidders per procurement lot.
- Capture key comparable fields: bidder name, bid amount, and submission time.
- Export structured output for manual review and further quantitative analysis.

### How it works (short)

- Scans each tender folder in `downloads/`.
- Prioritizes protocol sources (`PROPFIS`, `OPNPRT`, then other protocol files).
- Extracts tables from DOCX and PDF (with optional LLM fallback for hard PDFs).
- Merges duplicates and writes:
  - `bid_level` sheet (row-level records)
  - `tender_summary` sheet (procurement-level status)

### Security note

No API key is stored in source code. If LLM fallback is needed for PDF parsing,
set `OPENAI_API_KEY` in your environment before running the script.

### Example

```bash
python extract_bidders.py --downloads downloads --inventory inventory.xlsx --output bidder_extraction.xlsx
```

## Data Analysis

Compare your scraped data with the reference database:

### Quick Start

1. Install analysis dependencies:
   ```bash
   pip install -r requirements-analysis.txt
   ```

2. Run the analysis:
   ```bash
   python run_analysis.py
   ```

The analysis will:
- Compare scraped data with reference 2024 procurement IDs
- Focus on middle part of 2024 (April-September) to avoid year boundary issues
- Map downloaded files to procurement IDs
- Generate detailed reports and CSV exports

See [ANALYSIS_GUIDE.md](ANALYSIS_GUIDE.md) for detailed documentation.

## Initial failure and lesson learned

**Initial failure:**  
The first version only searched for `onclick="viewDocument(...)"` in the HTML.
On EIS pages, the document list is often *not* rendered directly into the HTML
rows, so no documents were found.

**Lesson learned:**  
For EIS, important data is frequently embedded in JavaScript arrays like
`ActualDocuments_items` and `ViewDocumentModel_Files_items`.  
The scraper now parses those arrays and uses the hidden JSON parameters to
build the correct download URLs.
