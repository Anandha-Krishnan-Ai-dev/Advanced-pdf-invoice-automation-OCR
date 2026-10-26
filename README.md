# Production PDF Invoice Extraction & Verification Pipeline

An automated Python pipeline that extracts structured invoice data from text-based and scanned PDFs, verifies invoice arithmetic using exact decimal math, detects duplicate invoices across historical runs via SQLite, and generates a formatted Excel audit report.

---

## Key Features

1. **Multi-Layout Regex Extraction:** Supports diverse invoice templates (`Invoice No`, `Tax Invoice #`, `Bill No`, `Customer`, `Bill To`, `Buyer`, `Grand Total`, `Amount Due`, and multi-line customer blocks).
2. **Automatic OCR Fallback:** Uses `pypdf` for fast embedded text extraction and automatically falls back to `PyMuPDF` (`fitz`) + `Tesseract OCR` (at 300 DPI) on a per-page basis when scanned pages are detected.
3. **Exact `Decimal` Math Verification:** Parses line items, taxes (`CGST`, `SGST`, `IGST`, `VAT`), and discounts using `decimal.Decimal` (`ROUND_HALF_UP`). Automatically flags invoices where `Line Items + Tax - Discount != Total Amount`.
4. **Cross-Run Duplicate Detection:** Persists SHA-256 file hashes and normalized invoice numbers in a local SQLite database (`output/invoice_history.db`) to flag duplicates both within the current batch and across prior runs.
5. **Production Safeguards:**
   - **PII Log Redaction:** Automatically masks customer names, email addresses, phone numbers, and tax IDs (GSTIN) in console and rotating log files.
   - **Automated File Routing:** Moves processed PDFs out of `input/` into `output/archive/` (passed), `output/review_queue/` (flagged for human review), or `output/quarantine/` (corrupt/unreadable).
   - **Atomic Excel Writes:** Writes to a temporary workbook before replacing `extracted_invoices.xlsx` to prevent file corruption.
   - **Built-in Unit Test Suite:** Includes a zero-dependency `--self-test` mode.

---

## Directory Structure

```text
.
├── extract_invoices.py          # Main automation script
├── README.md                    # Project documentation
├── input/                       # Drop incoming PDF invoices here
└── output/
    ├── extracted_invoices.xlsx  # Formatted Excel report (auto-filtered & color-coded)
    ├── invoice_history.db       # SQLite database for cross-run deduplication
    ├── automation.log           # Rotating log file (2 MB max x 5 backups)
    ├── text/                    # Raw extracted text (.txt) per PDF for auditing
    ├── archive/                 # Clean PDFs that passed all checks (Status: OK)
    ├── review_queue/            # PDFs flagged for manual verification (Status: REVIEW)
    └── quarantine/              # Unreadable, encrypted, or corrupt PDFs