# PDF Invoice Expense Extraction

A robust Python pipeline for extracting, merging, and summarizing expense data from scanned PDF invoices and exporting clean CSVs.  
All supplier names and business-specific logic are generalized for public sharing.

## Features

- Extracts expense values and dates from PDF invoices using rule-based parsing
- Handles batch processing for folders of invoices
- Merges and summarizes CSVs for expense analysis
- Ignores duplicates and allows specific override logic per supplier
- All supplier and customer info are anonymized or removed

## Usage

1. Install requirements:  
pip install -r requirements.txt

2. Run the scripts in `src/` as required (see each script for parameters).

## Security

- No sensitive or proprietary values included.
- All business/supplier names and company details are now placeholders.

---

## File Descriptions

- `src/pdf_extract_main.py` — Legacy batch extraction logic with error tracking.
- `src/pdf_folder_summary.py` — Counts pages/files and summarizes folder size.
- `src/pdf_summary_final.py` — Aggregates final expenses per supplier from merged CSVs.
- `src/pdf_merge_csvs.py` — Merges CSVs with the same base name into one.
- `src/pdf_expense_extractor.py` — Main, modernized, rule-based PDF-to-CSV extractor.