# Latvian EIS Tender Tools

This repository contains **two programs**:

1. **EIS Document Downloader** (`src/eis_scraper/`)  
   Downloads procurement documents and builds an inventory.
2. **Bidder Extractor** (`extract_bidders.py`)  
   Extracts bidder rows (company, lot, amount, timestamp) from opening protocols.

## Which one to run

- If you need to collect files from EIS, run the downloader.
- If you already have downloaded files and need bidder-level output, run the extractor.

## Program 1: EIS Document Downloader

Purpose: build a local document inventory from public EIS procurements.

Install:
```bash
python -m pip install -e .
```

Run:
```bash
python -m eis_scraper.cli --input eis_e_iepirkumi_izsludinatie_2026.csv
```

Small test:
```bash
python -m eis_scraper.cli --limit 2
```

Output:
- `downloads/<procurement_id>/...`
- `outputs/inventory.csv`
- `outputs/inventory.xlsx`
- `logs/errors.jsonl`

## Program 2: Bidder Extractor

Purpose: produce structured bidder-level data from protocol documents.

Install dependencies:
```bash
pip install -r requirements-extract-bidders.txt
```

Run:
```bash
python extract_bidders.py --downloads downloads --inventory inventory.xlsx --output bidder_extraction.xlsx
```

Output sheets:
- `bid_level`
- `tender_summary`

Security:
- No API key is stored in source code.
- For optional PDF LLM fallback, set `OPENAI_API_KEY` in your environment.

## Notes

- `run_analysis.py` and `ANALYSIS_GUIDE.md` are for validation/reporting of results.
