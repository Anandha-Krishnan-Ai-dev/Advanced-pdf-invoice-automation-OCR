import argparse
import hashlib
import logging
import os
import re
import shutil
import sqlite3
import sys
import tempfile
import unittest
from collections import Counter
from datetime import datetime, timezone
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from io import BytesIO
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Any, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from pypdf import PdfReader

# Optional OCR imports (graceful fallback if not installed)
try:
    import fitz  # PyMuPDF
    import pytesseract
    from PIL import Image

    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# ============================================================
# CONFIGURATION
# ============================================================
BASE_DIR = Path(__file__).resolve().parent
INPUT_FOLDER = BASE_DIR / "input"
OUTPUT_FOLDER = BASE_DIR / "output"
TEXT_FOLDER = OUTPUT_FOLDER / "text"
ARCHIVE_FOLDER = OUTPUT_FOLDER / "archive"
REVIEW_FOLDER = OUTPUT_FOLDER / "review_queue"
QUARANTINE_FOLDER = OUTPUT_FOLDER / "quarantine"

EXCEL_FILE = OUTPUT_FOLDER / "extracted_invoices.xlsx"
DB_FILE = OUTPUT_FOLDER / "invoice_history.db"
LOG_FILE = OUTPUT_FOLDER / "automation.log"

MAX_COLUMN_WIDTH = 50
MAX_LOG_BYTES = 2 * 1024 * 1024
BACKUP_COUNT = 5
MIN_TEXT_LENGTH_PER_PAGE = 20
OCR_DPI = 300

CURRENCY_FORMATS = {
    "INR": '"₹"#,##0.00;[Red]-"₹"#,##0.00',
    "USD": '"$"#,##0.00;[Red]-"$"#,##0.00',
    "EUR": '"€"#,##0.00;[Red]-"€"#,##0.00',
    "GBP": '"£"#,##0.00;[Red]-"£"#,##0.00',
}

CURRENCY_SYMBOL_MAP = {
    "₹": "INR",
    "INR": "INR",
    "RS": "INR",
    "RS.": "INR",
    "$": "USD",
    "USD": "USD",
    "€": "EUR",
    "EUR": "EUR",
    "£": "GBP",
    "GBP": "GBP",
}

HEADERS = [
    "Invoice Number",
    "Customer",
    "Description",
    "Currency",
    "Line Items Sum",
    "Tax Amount",
    "Discount Amount",
    "Total Amount",
    "Extraction Method",
    "Source PDF",
    "SHA-256 (Short)",
    "Status",
    "Review Notes",
]

# ============================================================
# UPGRADE 1: MULTI-LAYOUT REGEX FALLBACKS
# ============================================================
CURRENCY_TOKEN = r"(?:INR|USD|EUR|GBP|Rs\.?|₹|\$|€|£)"
AMOUNT_TOKEN = r"([+-]?\d[\d,]*(?:\.\d{1,2})?)"

INVOICE_NUMBER_PATTERNS = [
    re.compile(
        r"\b(?:Tax\s+)?Invoice\s*(?:Number|No\.?|#|ID|Code)\s*[:#\-]?\s*"
        r"([A-Za-z0-9][A-Za-z0-9/_-]{1,35})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\b(?:Bill|Document|Ref(?:erence)?)\s*(?:Number|No\.?|#)\s*[:#\-]?\s*"
        r"([A-Za-z0-9][A-Za-z0-9/_-]{1,35})",
        re.IGNORECASE,
    ),
    re.compile(
        r"\bInv\s*[:#\-]\s*([A-Za-z0-9][A-Za-z0-9/_-]{1,35})",
        re.IGNORECASE,
    ),
]

CUSTOMER_PATTERNS = [
    # Single-line label: "Customer: Acme Corp" or "Bill To: Acme Corp"
    re.compile(
        r"^\s*(?:Customer(?:\s+Name)?|Bill(?:ed)?\s+To|Client(?:\s+Name)?|Buyer)\s*[:\-]\s*(.+?)\s*$",
        re.IGNORECASE | re.MULTILINE,
    ),
    # Multi-line block where customer name is on the line immediately after "Bill To:"
    re.compile(
        r"^\s*(?:Customer|Bill(?:ed)?\s+To|Client|Buyer)\s*[:\-]?\s*\r?\n\s*([A-Za-z0-9][^\r\n]{2,80}?)\s*$",
        re.IGNORECASE | re.MULTILINE,
    ),
]

DESCRIPTION_BLOCK_PATTERNS = [
    re.compile(
        r"\b(?:Description|Particulars|Item(?:\s+Description)?|Service)\s+"
        r"(?:Qty\s+|Quantity\s+|Rate\s+|Price\s+)*Amount\b"
        r"(.*?)"
        r"\b(?:Subtotal|Sub-Total|Tax|CGST|SGST|IGST|GST|VAT|Discount|Total\s+Amount|Grand\s+Total|Amount\s+Due)\b",
        re.IGNORECASE | re.DOTALL,
    ),
]

TOTAL_AMOUNT_PATTERNS = [
    re.compile(
        r"\b(?:Grand\s+Total|Total\s+Amount|Amount\s+Due|Balance\s+Due|Net\s+Payable|Invoice\s+Total)\s*[:\-]?\s*"
        rf"({CURRENCY_TOKEN})?\s*{AMOUNT_TOKEN}",
        re.IGNORECASE,
    ),
    re.compile(
        r"^\s*Total\s*[:\-]?\s*"
        rf"({CURRENCY_TOKEN})?\s*{AMOUNT_TOKEN}\s*$",
        re.IGNORECASE | re.MULTILINE,
    ),
]

TAX_LINE_PATTERN = re.compile(
    r"^\s*(?:Tax|GST|IGST|CGST|SGST|VAT|Sales\s+Tax)(?:\s*\(\d+(?:\.\d+)?%\))?\s*[:\-]?\s*"
    rf"(?:{CURRENCY_TOKEN}\s*)?{AMOUNT_TOKEN}\s*$",
    re.IGNORECASE | re.MULTILINE,
)

DISCOUNT_LINE_PATTERN = re.compile(
    r"^\s*(?:Less\s*:?\s*)?Discount(?:\s*\(\d+(?:\.\d+)?%\))?\s*[:\-]?\s*"
    rf"[-+]?\s*(?:{CURRENCY_TOKEN}\s*)?{AMOUNT_TOKEN}\s*$",
    re.IGNORECASE | re.MULTILINE,
)

LINE_ITEM_AMOUNT_PATTERN = re.compile(
    rf"^(?P<desc>.+?)\s+(?:{CURRENCY_TOKEN}\s*)?(?P<amt>[+-]?\d[\d,]*(?:\.\d{{1,2}})?)\s*$"
)


# ============================================================
# UPGRADE 5A: PII REDACTION IN LOGS
# ============================================================
class PiiRedactingFilter(logging.Filter):
    """Redact emails, phone numbers, tax IDs, and tagged customer names in logs."""

    EMAIL_RE = re.compile(r"\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b")
    PHONE_RE = re.compile(r"\b(?:\+?\d{1,3}[-.\s]?)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}\b")
    GSTIN_RE = re.compile(r"\b\d{2}[A-Z]{5}\d{4}[A-Z][1-9A-Z]Z[0-9A-Z]\b")

    def __init__(self, enabled: bool = True) -> None:
        super().__init__()
        self.enabled = enabled

    def filter(self, record: logging.LogRecord) -> bool:
        if not self.enabled:
            return True
        msg = record.getMessage()
        msg = self.EMAIL_RE.sub("[REDACTED_EMAIL]", msg)
        msg = self.PHONE_RE.sub("[REDACTED_PHONE]", msg)
        msg = self.GSTIN_RE.sub("[REDACTED_TAX_ID]", msg)
        record.msg = msg
        record.args = ()
        return True


def mask_customer_name(name: Optional[str], redact: bool = True) -> str:
    """Mask customer name for safe logging (e.g., 'Acme Corp' -> 'Ac***p')."""
    if not name:
        return "Not Found"
    if not redact or len(name) <= 3:
        return name
    return f"{name[:2]}***{name[-1]}"


# ============================================================
# DIRECTORY AND LOGGING SETUP
# ============================================================
def setup_directories() -> None:
    """Create all required working and lifecycle directories."""
    for folder in (
        INPUT_FOLDER,
        OUTPUT_FOLDER,
        TEXT_FOLDER,
        ARCHIVE_FOLDER,
        REVIEW_FOLDER,
        QUARANTINE_FOLDER,
    ):
        folder.mkdir(parents=True, exist_ok=True)


def setup_logging(redact_pii: bool = True) -> None:
    """Log to both the console and a rotating log file with optional PII redaction."""
    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    pii_filter = PiiRedactingFilter(enabled=redact_pii)

    file_handler = RotatingFileHandler(
        LOG_FILE,
        maxBytes=MAX_LOG_BYTES,
        backupCount=BACKUP_COUNT,
        encoding="utf-8",
    )
    file_handler.setFormatter(formatter)
    file_handler.addFilter(pii_filter)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    console_handler.addFilter(pii_filter)

    logging.basicConfig(
        level=logging.INFO,
        handlers=[file_handler, console_handler],
        force=True,
    )


# ============================================================
# UPGRADE 2: PDF TEXT EXTRACTION WITH OCR FALLBACK
# ============================================================
def ocr_pdf_page(pdf_path: Path, page_index: int) -> str:
    """Render a single PDF page to an image and run Tesseract OCR."""
    if not OCR_AVAILABLE:
        raise RuntimeError(
            "OCR libraries (pymupdf, pytesseract, Pillow) are not installed."
        )

    with fitz.open(str(pdf_path)) as doc:
        page = doc.load_page(page_index)
        pix = page.get_pixmap(dpi=OCR_DPI)
        img_bytes = pix.tobytes("png")

    with Image.open(BytesIO(img_bytes)) as img:
        text = pytesseract.image_to_string(img)
    return text.strip()


def extract_pdf_text(pdf_path: Path) -> tuple[str, str, list[str]]:
    """
    Extract text from all pages of a PDF, falling back to OCR per page if needed.
    Returns: (extracted_text, extraction_method, warnings_list)
    """
    reader = PdfReader(str(pdf_path), strict=False)
    warnings: list[str] = []

    if reader.is_encrypted:
        try:
            result = reader.decrypt("")
            if not result:
                raise ValueError("PDF is password-protected.")
        except Exception as exc:
            raise ValueError(
                "Unable to decrypt PDF. Check its password."
            ) from exc

    pages: list[str] = []
    used_text = False
    used_ocr = False

    for page_index, page in enumerate(reader.pages):
        page_number = page_index + 1
        page_text = ""
        try:
            page_text = (page.extract_text() or "").strip()
        except Exception as exc:
            logging.warning(
                "pypdf could not extract page %d from %s: %s",
                page_number,
                pdf_path.name,
                exc,
            )

        if len(page_text) >= MIN_TEXT_LENGTH_PER_PAGE:
            pages.append(page_text)
            used_text = True
            continue

        # Fallback to OCR when embedded text is missing or too short
        try:
            ocr_text = ocr_pdf_page(pdf_path, page_index)
            if ocr_text:
                pages.append(ocr_text)
                used_ocr = True
                logging.info(
                    "Page %d of %s extracted via OCR.",
                    page_number,
                    pdf_path.name,
                )
            elif page_text:
                pages.append(page_text)
                used_text = True
        except Exception as exc:
            msg = f"OCR fallback unavailable on page {page_number}: {exc}"
            logging.warning("%s (%s)", msg, pdf_path.name)
            warnings.append(msg)
            if page_text:
                pages.append(page_text)
                used_text = True

    if used_text and used_ocr:
        method = "HYBRID"
    elif used_ocr:
        method = "OCR"
    elif used_text:
        method = "TEXT"
    else:
        method = "NONE"

    return "\n".join(pages).strip(), method, warnings


def save_text_file(pdf_path: Path, text: str) -> Path:
    """Save extracted text for debugging and audit purposes."""
    output_path = TEXT_FOLDER / f"{pdf_path.stem}.txt"
    output_path.write_text(text, encoding="utf-8")
    return output_path


def compute_file_sha256(file_path: Path) -> str:
    """Compute the SHA-256 digest of a file for exact duplicate detection."""
    hasher = hashlib.sha256()
    with file_path.open("rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


# ============================================================
# UPGRADE 1 & 3: MULTI-LAYOUT EXTRACTION & DECIMAL ARITHMETIC
# ============================================================
def parse_decimal(raw_value: Optional[str]) -> Optional[Decimal]:
    """Convert a string amount into a 2-decimal-place Decimal safely."""
    if not raw_value:
        return None
    cleaned = raw_value.replace(",", "").strip()
    try:
        value = Decimal(cleaned)
    except InvalidOperation:
        return None
    if not value.is_finite():
        return None
    return value.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def normalize_currency(raw_symbol: Optional[str]) -> str:
    """Normalize a currency symbol or code to an ISO 4217 code (default INR)."""
    if not raw_symbol:
        return "INR"
    key = raw_symbol.strip().upper()
    return CURRENCY_SYMBOL_MAP.get(key, "INR")


def extract_invoice_number(text: str) -> Optional[str]:
    """Try multiple invoice number regex patterns."""
    for pattern in INVOICE_NUMBER_PATTERNS:
        match = pattern.search(text)
        if match:
            candidate = match.group(1).strip()
            if candidate.lower() not in {"date", "number", "no", "total"}:
                return candidate
    return None


def extract_customer(text: str) -> Optional[str]:
    """Try multiple customer/Bill-To regex patterns."""
    for pattern in CUSTOMER_PATTERNS:
        match = pattern.search(text)
        if match:
            value = match.group(1).strip()
            if value:
                return value
    return None


def extract_total_and_currency(text: str) -> tuple[Optional[Decimal], str]:
    """Extract the total amount as a Decimal along with the detected currency."""
    for pattern in TOTAL_AMOUNT_PATTERNS:
        matches = list(pattern.finditer(text))
        if matches:
            # Prefer the last match in case "Total" appears in headers
            last_match = matches[-1]
            raw_curr, raw_amt = last_match.group(1), last_match.group(2)
            amount = parse_decimal(raw_amt)
            if amount is not None:
                return amount, normalize_currency(raw_curr)
    return None, "INR"


def extract_sum_of_matches(pattern: re.Pattern, text: str) -> Optional[Decimal]:
    """Sum all occurrences of a pattern (e.g., CGST + SGST lines)."""
    matches = pattern.findall(text)
    if not matches:
        return None
    total = Decimal("0.00")
    found_valid = False
    for raw_amt in matches:
        val = parse_decimal(raw_amt)
        if val is not None:
            total += val
            found_valid = True
    return total if found_valid else None


def extract_description_and_line_items(
    text: str,
) -> tuple[Optional[str], Optional[Decimal]]:
    """
    Extract both the human-readable description and the sum of line-item amounts.
    """
    for pattern in DESCRIPTION_BLOCK_PATTERNS:
        match = pattern.search(text)
        if not match:
            continue

        raw_block = match.group(1).strip()
        if not raw_block:
            continue

        descriptions: list[str] = []
        line_sum = Decimal("0.00")
        has_line_amounts = False

        for raw_line in raw_block.splitlines():
            line = raw_line.strip()
            if not line:
                continue

            item_match = LINE_ITEM_AMOUNT_PATTERN.match(line)
            if item_match:
                desc_part = item_match.group("desc").strip()
                amt_val = parse_decimal(item_match.group("amt"))
                if amt_val is not None:
                    line_sum += amt_val
                    has_line_amounts = True
                if desc_part:
                    descriptions.append(desc_part)
            else:
                descriptions.append(line)

        cleaned_desc = re.sub(r"\s+", " ", "; ".join(descriptions)).strip()
        return (
            cleaned_desc or None,
            line_sum if has_line_amounts else None,
        )

    return None, None


# ============================================================
# VALIDATION AND ARITHMETIC VERIFICATION
# ============================================================
def build_record(
    pdf_file: Path,
    file_hash: str,
    text: str,
    extraction_method: str,
    extraction_warnings: list[str],
) -> dict[str, Any]:
    """Extract fields, verify Decimal arithmetic, and flag records needing review."""
    invoice_number = extract_invoice_number(text)
    customer = extract_customer(text)
    description, line_items_sum = extract_description_and_line_items(text)
    total_amount, currency = extract_total_and_currency(text)
    tax_amount = extract_sum_of_matches(TAX_LINE_PATTERN, text)
    discount_amount = extract_sum_of_matches(DISCOUNT_LINE_PATTERN, text)

    review_notes: list[str] = list(extraction_warnings)

    if extraction_method in {"OCR", "HYBRID"}:
        review_notes.append(f"Extracted via {extraction_method}; verify OCR accuracy")

    record: dict[str, Any] = {
        "invoice_number": invoice_number,
        "customer": customer,
        "description": description,
        "currency": currency,
        "line_items_sum": line_items_sum,
        "tax_amount": tax_amount,
        "discount_amount": discount_amount,
        "total_amount": total_amount,
        "extraction_method": extraction_method,
        "source_pdf": pdf_file.name,
        "file_hash": file_hash,
        "status": "OK",
        "review_notes": review_notes,
    }

    required_fields = {
        "invoice_number": "Invoice number not found",
        "customer": "Customer not found",
        "description": "Description not found",
        "total_amount": "Total amount not found or invalid",
    }

    for field, message in required_fields.items():
        if record[field] is None:
            record["review_notes"].append(message)

    if total_amount is not None and total_amount <= Decimal("0.00"):
        record["review_notes"].append(
            f"Non-positive total amount ({total_amount}); verify credit note or error"
        )

    # Independent arithmetic verification when line items are present
    if line_items_sum is not None and total_amount is not None:
        tax = tax_amount or Decimal("0.00")
        discount = discount_amount or Decimal("0.00")
        expected_total = (line_items_sum + tax - discount).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        if expected_total != total_amount:
            record["review_notes"].append(
                f"Arithmetic mismatch: items ({line_items_sum}) + tax ({tax}) "
                f"- discount ({discount}) = {expected_total} != total ({total_amount})"
            )

    if record["review_notes"]:
        record["status"] = "REVIEW"

    return record


# ============================================================
# UPGRADE 4: CROSS-RUN DUPLICATE DETECTION (SQLITE)
# ============================================================
def init_history_db(db_path: Path) -> None:
    """Create the SQLite table for tracking processed invoices across runs."""
    with sqlite3.connect(db_path) as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS processed_invoices (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_hash TEXT NOT NULL,
                invoice_number_norm TEXT,
                customer TEXT,
                total_amount TEXT,
                currency TEXT,
                source_pdf TEXT NOT NULL,
                status TEXT NOT NULL,
                processed_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_file_hash ON processed_invoices(file_hash)"
        )
        conn.execute(
            "CREATE INDEX IF NOT EXISTS idx_inv_norm ON processed_invoices(invoice_number_norm)"
        )


def flag_duplicates(records: list[dict[str, Any]], db_path: Path) -> None:
    """Flag both in-batch and cross-run duplicates using SQLite history."""
    # 1. In-batch duplicate invoice numbers and file hashes
    inv_counts = Counter(
        r["invoice_number"].casefold()
        for r in records
        if r["invoice_number"]
    )
    hash_counts = Counter(r["file_hash"] for r in records)

    for record in records:
        inv = record["invoice_number"]
        if inv and inv_counts[inv.casefold()] > 1:
            record["review_notes"].append(
                "Duplicate invoice number within current batch"
            )
            record["status"] = "REVIEW"

        if hash_counts[record["file_hash"]] > 1:
            record["review_notes"].append(
                "Identical PDF file (SHA-256) repeated in current batch"
            )
            record["status"] = "REVIEW"

    # 2. Historical cross-run duplicate check via SQLite
    with sqlite3.connect(db_path) as conn:
        cursor = conn.cursor()
        for record in records:
            cursor.execute(
                "SELECT source_pdf, processed_at FROM processed_invoices WHERE file_hash = ? LIMIT 1",
                (record["file_hash"],),
            )
            hash_row = cursor.fetchone()
            if hash_row:
                record["review_notes"].append(
                    f"File hash already processed in prior run ({hash_row[0]} at {hash_row[1]})"
                )
                record["status"] = "REVIEW"

            inv = record["invoice_number"]
            if inv:
                cursor.execute(
                    "SELECT source_pdf, processed_at FROM processed_invoices "
                    "WHERE invoice_number_norm = ? AND file_hash != ? LIMIT 1",
                    (inv.casefold(), record["file_hash"]),
                )
                inv_row = cursor.fetchone()
                if inv_row:
                    record["review_notes"].append(
                        f"Invoice number already processed in prior run ({inv_row[0]} at {inv_row[1]})"
                    )
                    record["status"] = "REVIEW"


def record_run_in_history_db(
    records: list[dict[str, Any]], db_path: Path
) -> None:
    """Persist processed records into SQLite after Excel generation succeeds."""
    now_iso = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M:%S UTC")
    rows = [
        (
            r["file_hash"],
            r["invoice_number"].casefold() if r["invoice_number"] else None,
            r["customer"],
            str(r["total_amount"]) if r["total_amount"] is not None else None,
            r["currency"],
            r["source_pdf"],
            r["status"],
            now_iso,
        )
        for r in records
    ]
    with sqlite3.connect(db_path) as conn:
        conn.executemany(
            """
            INSERT INTO processed_invoices (
                file_hash, invoice_number_norm, customer,
                total_amount, currency, source_pdf, status, processed_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            rows,
        )


# ============================================================
# UPGRADE 5B: SAFE FILE ROUTING (ARCHIVE / REVIEW / QUARANTINE)
# ============================================================
def move_file_safely(src: Path, dest_folder: Path) -> Path:
    """Move a file to dest_folder, appending a timestamp if a name collision exists."""
    dest_folder.mkdir(parents=True, exist_ok=True)
    target = dest_folder / src.name
    if target.exists():
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        target = dest_folder / f"{src.stem}_{timestamp}{src.suffix}"
    shutil.move(str(src), str(target))
    return target


# ============================================================
# EXCEL REPORT GENERATION
# ============================================================
def create_excel_report(
    records: list[dict[str, Any]],
    output_path: Path,
) -> None:
    """Create a formatted invoice report and save it atomically."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Invoice Report"

    ws.append(HEADERS)

    header_fill = PatternFill(fill_type="solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    review_fill = PatternFill(fill_type="solid", fgColor="FFF2CC")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    for record in records:
        ws.append([
            record["invoice_number"] or "Not Found",
            record["customer"] or "Not Found",
            record["description"] or "Not Found",
            record["currency"],
            float(record["line_items_sum"]) if record["line_items_sum"] is not None else None,
            float(record["tax_amount"]) if record["tax_amount"] is not None else None,
            float(record["discount_amount"]) if record["discount_amount"] is not None else None,
            float(record["total_amount"]) if record["total_amount"] is not None else None,
            record["extraction_method"],
            record["source_pdf"],
            record["file_hash"][:12],
            record["status"],
            "; ".join(record["review_notes"]),
        ])

        row_number = ws.max_row
        num_format = CURRENCY_FORMATS.get(
            record["currency"], CURRENCY_FORMATS["INR"]
        )

        # Format monetary columns (columns 5, 6, 7, 8)
        for col_idx in (5, 6, 7, 8):
            amount_cell = ws.cell(row=row_number, column=col_idx)
            if amount_cell.value is not None:
                amount_cell.number_format = num_format

        if record["status"] == "REVIEW":
            for cell in ws[row_number]:
                cell.fill = review_fill

    for column in ws.columns:
        max_length = max(
            (
                len(line)
                for cell in column
                if cell.value is not None
                for line in str(cell.value).splitlines()
            ),
            default=0,
        )
        column_letter = get_column_letter(column[0].column)
        ws.column_dimensions[column_letter].width = min(
            max_length + 3,
            MAX_COLUMN_WIDTH,
        )

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(
                vertical="top",
                wrap_text=cell.column in (3, 13),
            )

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    ws.row_dimensions[1].height = 30

    temp_path = output_path.with_name(f".{output_path.stem}.tmp.xlsx")
    try:
        wb.save(temp_path)
        os.replace(temp_path, output_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


# ============================================================
# UPGRADE 5C: BUILT-IN SELF-TEST SUITE (--self-test)
# ============================================================
class InvoiceAutomationTests(unittest.TestCase):
    """Automated unit tests verifying regex layouts, Decimal math, and SQLite dedup."""

    def test_multi_layout_extraction_and_valid_math(self) -> None:
        sample_text = (
            "Tax Invoice #\n"
            "Bill No: INV-2026-089\n"
            "Billed To:\n"
            "Acme Global Solutions Pvt Ltd\n"
            "Particulars Qty Rate Amount\n"
            "Cloud Hosting Services ₹ 10,000.00\n"
            "Managed Support ₹ 5,000.00\n"
            "CGST (9%) ₹ 1,350.00\n"
            "SGST (9%) ₹ 1,350.00\n"
            "Discount ₹ 700.00\n"
            "Grand Total: ₹ 17,000.00\n"
        )
        record = build_record(
            Path("sample_ok.pdf"),
            "a" * 64,
            sample_text,
            "TEXT",
            [],
        )
        self.assertEqual(record["invoice_number"], "INV-2026-089")
        self.assertEqual(record["customer"], "Acme Global Solutions Pvt Ltd")
        self.assertEqual(record["line_items_sum"], Decimal("15000.00"))
        self.assertEqual(record["tax_amount"], Decimal("2700.00"))
        self.assertEqual(record["discount_amount"], Decimal("700.00"))
        self.assertEqual(record["total_amount"], Decimal("17000.00"))
        self.assertEqual(record["status"], "OK")

    def test_arithmetic_mismatch_flagged(self) -> None:
        bad_math_text = (
            "Invoice Number: INV-900\n"
            "Customer: Beta Retail\n"
            "Description Amount\n"
            "Annual License $ 500.00\n"
            "Tax $ 50.00\n"
            "Total Amount: $ 600.00\n"
        )
        record = build_record(
            Path("bad_math.pdf"),
            "b" * 64,
            bad_math_text,
            "TEXT",
            [],
        )
        self.assertEqual(record["currency"], "USD")
        self.assertEqual(record["status"], "REVIEW")
        self.assertTrue(
            any("Arithmetic mismatch" in note for note in record["review_notes"])
        )

    def test_cross_run_sqlite_deduplication(self) -> None:
        with tempfile.TemporaryDirectory() as tmpdir:
            db_path = Path(tmpdir) / "test_history.db"
            init_history_db(db_path)

            first_batch = [
                {
                    "invoice_number": "INV-101",
                    "customer": "Gamma Corp",
                    "total_amount": Decimal("100.00"),
                    "currency": "INR",
                    "source_pdf": "batch1.pdf",
                    "file_hash": "1111" * 16,
                    "status": "OK",
                    "review_notes": [],
                }
            ]
            flag_duplicates(first_batch, db_path)
            self.assertEqual(first_batch[0]["status"], "OK")
            record_run_in_history_db(first_batch, db_path)

            second_batch = [
                {
                    "invoice_number": "inv-101",
                    "customer": "Gamma Corp",
                    "total_amount": Decimal("100.00"),
                    "currency": "INR",
                    "source_pdf": "batch2_reupload.pdf",
                    "file_hash": "2222" * 16,
                    "status": "OK",
                    "review_notes": [],
                }
            ]
            flag_duplicates(second_batch, db_path)
            self.assertEqual(second_batch[0]["status"], "REVIEW")
            self.assertTrue(
                any("prior run" in note for note in second_batch[0]["review_notes"])
            )


# ============================================================
# MAIN AUTOMATION
# ============================================================
def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Production PDF Invoice Extraction & Verification Pipeline"
    )
    parser.add_argument(
        "--self-test",
        action="store_true",
        help="Run the built-in unit test suite and exit.",
    )
    parser.add_argument(
        "--keep-input",
        action="store_true",
        help="Do not move PDFs out of input/ into archive/review/quarantine folders.",
    )
    parser.add_argument(
        "--no-redact",
        action="store_true",
        help="Disable PII redaction in console and log file output.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    if args.self_test:
        suite = unittest.defaultTestLoader.loadTestsFromTestCase(
            InvoiceAutomationTests
        )
        result = unittest.TextTestRunner(verbosity=2).run(suite)
        return 0 if result.wasSuccessful() else 1

    redact_pii = not args.no_redact
    setup_directories()
    setup_logging(redact_pii=redact_pii)
    init_history_db(DB_FILE)

    logging.info("=" * 60)
    logging.info(
        "PDF INVOICE AUTOMATION STARTED (OCR Available: %s)", OCR_AVAILABLE
    )
    logging.info("=" * 60)

    pdf_files = sorted(
        (
            path
            for path in INPUT_FOLDER.iterdir()
            if path.is_file() and path.suffix.lower() == ".pdf"
        ),
        key=lambda path: path.name.casefold(),
    )

    if not pdf_files:
        logging.warning("No PDF files found in %s", INPUT_FOLDER.resolve())
        return 0

    records: list[dict[str, Any]] = []
    pdf_map: dict[str, Path] = {}
    failed = 0

    for pdf_file in pdf_files:
        logging.info("Processing: %s", pdf_file.name)
        try:
            file_hash = compute_file_sha256(pdf_file)
            text, method, warnings = extract_pdf_text(pdf_file)

            if not text:
                logging.warning(
                    "No extractable text in %s (even after OCR check). Moving to quarantine.",
                    pdf_file.name,
                )
                failed += 1
                if not args.keep_input:
                    move_file_safely(pdf_file, QUARANTINE_FOLDER)
                continue

            save_text_file(pdf_file, text)

            record = build_record(pdf_file, file_hash, text, method, warnings)
            records.append(record)
            pdf_map[pdf_file.name] = pdf_file

        except Exception:
            failed += 1
            logging.exception("Failed to process %s", pdf_file.name)
            if not args.keep_input and pdf_file.exists():
                move_file_safely(pdf_file, QUARANTINE_FOLDER)

    # Check both in-batch and cross-run SQLite duplicates
    flag_duplicates(records, DB_FILE)

    for record in records:
        if record["status"] == "REVIEW":
            logging.warning(
                "Review required for %s: %s",
                record["source_pdf"],
                "; ".join(record["review_notes"]),
            )
        else:
            logging.info(
                "Extracted invoice %s | Customer: %s | Amount: %s %.2f",
                record["invoice_number"],
                mask_customer_name(record["customer"], redact=redact_pii),
                record["currency"],
                record["total_amount"],
            )

    successful = sum(r["status"] == "OK" for r in records)
    review_count = sum(r["status"] == "REVIEW" for r in records)

    if records:
        try:
            create_excel_report(records, EXCEL_FILE)
            record_run_in_history_db(records, DB_FILE)
            logging.info("Excel report saved: %s", EXCEL_FILE.resolve())

            # Route processed PDFs only after report and DB commit succeed
            if not args.keep_input:
                for record in records:
                    src_pdf = pdf_map.get(record["source_pdf"])
                    if src_pdf and src_pdf.exists():
                        dest = (
                            ARCHIVE_FOLDER
                            if record["status"] == "OK"
                            else REVIEW_FOLDER
                        )
                        move_file_safely(src_pdf, dest)

        except PermissionError:
            logging.exception(
                "Cannot replace the Excel report. Close the workbook in Excel and retry."
            )
            return 1
        except Exception:
            logging.exception("Failed to create the Excel report.")
            return 1
    else:
        logging.warning(
            "No invoice records were extracted; no new report was generated."
        )

    logging.info("=" * 60)
    logging.info("FINAL REPORT")
    logging.info("PDF files found : %d", len(pdf_files))
    logging.info("Processed       : %d", len(records))
    logging.info("Passed checks   : %d", successful)
    logging.info("Needs review    : %d", review_count)
    logging.info("Failed/skipped  : %d", failed)
    logging.info("History DB      : %s", DB_FILE.resolve())
    logging.info("Log file        : %s", LOG_FILE.resolve())
    logging.info("=" * 60)

    return 1 if failed else 0


if __name__ == "__main__":
    raise SystemExit(main())