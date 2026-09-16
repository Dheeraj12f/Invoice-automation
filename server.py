"""MCP server for extracting invoice data into a local JSON file.

Claude reads invoice scans through MCP. This server stores the values Claude
extracts; it does not call an external AI API and does not use Excel.
"""

from __future__ import annotations

import base64
import json
import re
import shutil
import unicodedata
import xml.etree.ElementTree as ET
import zipfile
from datetime import date, datetime, timezone
from pathlib import Path

import fitz
import openpyxl
from mcp.server.mcpserver import MCPServer
from mcp.types import ImageContent, TextContent

mcp = MCPServer("Invoice Processing Server")

BASE_DIR = Path(__file__).resolve().parent
INVOICES_DIR = BASE_DIR / "invoices"
INBOX_DIR = INVOICES_DIR / "inbox"
PROCESSED_DIR = INVOICES_DIR / "processed"  # Successfully extracted files are moved here, never deleted.
FAILED_DIR = INVOICES_DIR / "failed"
EXTRACTED_DATA_PATH = BASE_DIR / "extracted_invoices.json"
RENAME_MAP_PATH = INVOICES_DIR / ".filename_renames.json"

for directory in (INBOX_DIR, PROCESSED_DIR, FAILED_DIR):
    directory.mkdir(parents=True, exist_ok=True)


def _is_lookalike_char(ch: str) -> bool:
    """True for a character that is invisible or renders indistinguishably from a
    normal space/character, but is byte-different - e.g. the narrow no-break space
    (U+202F) that Windows/Android insert before "am"/"pm" in scanner app filenames,
    or a zero-width joiner. These break exact filename matching between what an
    MCP client sees/types and what is actually on disk.
    """
    if ch == " ":
        return False
    category = unicodedata.category(ch)
    # Zs = all Unicode space separators (no-break space, narrow no-break space,
    # thin space, ideographic space, ...); Cf = invisible format chars (zero-width
    # space/joiner, byte-order mark, bidi marks, ...).
    return category in ("Zs", "Cf")


def sanitize_filename(name: str) -> str:
    """Replace lookalike/invisible characters with a plain space, then collapse
    and trim so the visible spelling of the filename is unchanged."""
    cleaned = "".join(" " if _is_lookalike_char(ch) else ch for ch in name)
    cleaned = re.sub(r" {2,}", " ", cleaned).strip()
    return cleaned or name


def load_rename_map() -> dict:
    if not RENAME_MAP_PATH.exists():
        return {}
    try:
        return json.loads(RENAME_MAP_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_rename_map(mapping: dict) -> None:
    RENAME_MAP_PATH.write_text(json.dumps(mapping, indent=2, ensure_ascii=False) + "\n", encoding="utf-8")


def sanitize_inbox_filenames() -> bool:
    """Rename any inbox file whose name contains a lookalike/invisible character
    to a plain-ASCII-space equivalent, so it can reliably be looked up by name
    later. The original name is preserved in RENAME_MAP_PATH so it can still be
    shown (in brackets) alongside the renamed one in extracted_invoices.json.
    Returns True if anything was renamed.
    """
    mapping = load_rename_map()
    changed = False
    for path in INBOX_DIR.iterdir():
        if not path.is_file() or path.name.startswith("~$"):
            continue
        clean_name = sanitize_filename(path.name)
        if clean_name == path.name:
            continue
        target = INBOX_DIR / clean_name
        counter = 1
        while target.exists():
            target = INBOX_DIR / f"{target.stem} ({counter}){target.suffix}"
            counter += 1
        try:
            path.rename(target)
        except Exception:
            continue
        mapping[target.name] = path.name
        changed = True
    if changed:
        save_rename_map(mapping)
    return changed


def display_filename(filename: str) -> str:
    """Original filename, with the on-disk renamed filename appended in brackets
    if this file was auto-renamed by sanitize_inbox_filenames(). Used only for
    the human-readable record written to extracted_invoices.json."""
    original = load_rename_map().get(filename)
    if original and original != filename:
        return f"{original} ({filename})"
    return filename


def inbox_file(filename: str) -> Path:
    """Return a direct file in inbox and reject path traversal."""
    path = (INBOX_DIR / filename).resolve()
    if path.parent != INBOX_DIR.resolve() or not path.is_file():
        # The exact name might not exist because it still has a lookalike/invisible
        # character (e.g. a client retrying with a plain space it typed itself, or
        # a new file dropped into inbox since the last list_inbox_invoices call).
        # Run a sanitize pass and retry once before giving up.
        if sanitize_inbox_filenames():
            path = (INBOX_DIR / filename).resolve()
            if path.parent == INBOX_DIR.resolve() and path.is_file():
                return path
        raise FileNotFoundError(f"File {filename} was not found in inbox.")
    return path


def load_json_payload() -> dict:
    if not EXTRACTED_DATA_PATH.exists():
        return {"generated_at": None, "invoices": []}
    return json.loads(EXTRACTED_DATA_PATH.read_text(encoding="utf-8"))


def save_json_payload(payload: dict) -> None:
    EXTRACTED_DATA_PATH.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")


@mcp.tool()
def list_inbox_invoices() -> list[str]:
    """List invoice PDFs, Word documents (.docx), Excel spreadsheets (.xlsx), and images waiting in the inbox."""
    sanitize_inbox_filenames()
    supported = {".pdf", ".png", ".jpg", ".jpeg", ".bmp", ".docx", ".doc", ".xlsx", ".xls"}
    return sorted(
        path.name for path in INBOX_DIR.iterdir()
        if path.is_file()
        and path.suffix.lower() in supported
        and not path.name.startswith("~$")
    )


def extract_docx_content(file_path: Path) -> str:
    """Extract all text, paragraphs, and tables from a Word (.docx) document."""
    try:
        with zipfile.ZipFile(file_path) as z:
            xml_bytes = z.read("word/document.xml")
            tree = ET.fromstring(xml_bytes)
            ns = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            lines = []
            for elem in tree.iter():
                if elem.tag == f"{ns}p":
                    para = "".join(t.text for t in elem.iter(f"{ns}t") if t.text).strip()
                    if para:
                        lines.append(para)
                elif elem.tag == f"{ns}tr":
                    row_cells = []
                    for tc in elem.iter(f"{ns}tc"):
                        c_text = " ".join(t.text for t in tc.iter(f"{ns}t") if t.text).strip()
                        row_cells.append(c_text)
                    if any(row_cells):
                        lines.append(" | ".join(row_cells))
            return "\n".join(lines)
    except Exception as e:
        # Fallback to PyMuPDF if zipfile parsing encounters an unexpected structure
        try:
            with fitz.open(file_path) as doc:
                return "\n".join(page.get_text() for page in doc).strip()
        except Exception:
            raise RuntimeError(f"Could not extract Word document content: {e}")


def extract_excel_content(file_path: Path, max_rows_per_sheet: int = 200) -> str:
    """Extract sheets, table headers, and rows from an Excel (.xlsx, .xls) document."""
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        lines = []
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            lines.append(f"=== Sheet: {sheet_name} ===")
            row_count = 0
            for row in sheet.iter_rows(values_only=True):
                if any(cell is not None and str(cell).strip() != "" for cell in row):
                    row_str = " | ".join(str(cell).strip() if cell is not None else "" for cell in row)
                    lines.append(row_str)
                    row_count += 1
                    if row_count >= max_rows_per_sheet:
                        lines.append(f"... (truncated at {max_rows_per_sheet} rows)")
                        break
        wb.close()
        return "\n".join(lines)
    except Exception as e:
        raise RuntimeError(f"Could not extract Excel spreadsheet content: {e}")


def optimize_image_payload(
    doc_or_page: fitz.Page | Path,
    max_dim: int = 1800,
    max_bytes: int = 900 * 1024,
) -> ImageContent:
    """Render and compress an image or PDF page to JPEG under 1MB for fast, reliable MCP transmission."""
    if isinstance(doc_or_page, Path):
        with fitz.open(doc_or_page) as doc:
            page = doc[0]
            w, h = page.rect.width, page.rect.height
            longest = max(w, h)
            scale = min(1.0, max_dim / longest) if longest > max_dim else 1.0
            if longest < 1200:
                scale = min(max_dim / longest, 2.0)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
            data = pix.tobytes("jpg")
            while len(data) > max_bytes and scale > 0.35:
                scale *= 0.8
                pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
                data = pix.tobytes("jpg")
    else:
        page = doc_or_page
        w, h = page.rect.width, page.rect.height
        longest = max(w, h)
        scale = min(1.0, max_dim / longest) if longest > max_dim else 1.0
        if longest < 1200:
            scale = min(max_dim / longest, 2.0)
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
        data = pix.tobytes("jpg")
        while len(data) > max_bytes and scale > 0.35:
            scale *= 0.8
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
            data = pix.tobytes("jpg")

    return ImageContent(
        data=base64.b64encode(data).decode("ascii"),
        mimeType="image/jpeg",
    )


INVOICE_PROMPT_BANNER = (
    "=== INVOICE EXTRACTION INSTRUCTIONS FOR CLAUDE AI ===\n"
    "Carefully analyze this invoice (description, line items, SAC code, header, tables) to extract:\n"
    "1. 'amount_ai': EXTRACT THE TOTAL TAXABLE VALUE (Taxable Amount / Base Amount before GST taxes).\n"
    "   - Look for the field labelled 'Total Taxable Value', 'Taxable Amount', 'Taxable Value', 'Basic Amount', or 'Subtotal'.\n"
    "   - Do NOT extract the Grand Total (inclusive of GST). The Amount column in the ledger must contain the Taxable Value.\n"
    "   - For example, if Taxable Value = 1,834,250 and Grand Total = 2,164,414, pass amount_ai=1834250.0.\n"
    "2. 'invoice_date_ai': FORMAT INVOICE DATE AS M/D/YYYY (e.g. '9/1/2026', '9/2/2026', '8/1/2026') without leading zeros.\n"
    "3. 'tds_name':\n"
    "   - 'RENT-2026' for premises rent, lease, or licence fee\n"
    "   - 'Contract-2026' for CAM, maintenance, AMC, parking, electricity, or contractor charges\n"
    "   - 'Rent-(Urd)' for rent from unregistered vendors\n"
    "   - 'COMMISSION' for store commission / brokerage fees\n"
    "4. 'service_name':\n"
    "   - 'Rent_997212' for premises rent (SAC 997212)\n"
    "   - 'Maintenance.' for Common Area Maintenance (CAM) / maintenance charges\n"
    "   - 'AMC Charges.' for Annual Maintenance Contracts (HVAC, Lift, DG, equipment)\n"
    "   - 'PARKING RENT' for parking space rental\n"
    "   - 'Electricity Chrg' for electricity / power backup charges\n"
    "   - 'Commission Charges-Store_997221' for store commission\n"
    "   - 'Rent-(Urd)' for unregistered vendor rental\n"
    "   - 'General Expenses' for general miscellaneous contractor expenses\n"
    "5. 'remark': Any invoice description, period (e.g. 'Rent for Aug 2026'), or notes.\n"
    "6. 'vendor_name': EXTRACT THE VENDOR / SELLER / SUPPLIER / LANDLORD NAME DIRECTLY FROM THE INVOICE DOCUMENT ONLY (from header, letterhead, Seller details, or Bill From). Do NOT use Sheet1; extract it directly from the invoice.\n"
    "You MUST provide these decided values when calling save_extracted_data_to_json.\n"
    "=====================================================\n\n"
)


@mcp.tool()
def read_invoice_content(filename: str, page_number: int = 1) -> dict | ImageContent | list:
    """Return invoice text or a native MCP image for Claude to inspect.
    Supports PDF documents, Word documents (.docx), Excel spreadsheets (.xlsx), and images.
    Files larger than 1MB are automatically optimized and compressed to ensure fast, reliable processing.

    Args:
        filename: Name of the invoice file in the inbox.
        page_number: Page number to inspect for multi-page PDFs (1-indexed, default: 1).
    """
    try:
        file_path = inbox_file(filename)
    except FileNotFoundError as error:
        return {"success": False, "error": str(error)}

    extension = file_path.suffix.lower()

    # Word documents (.docx, .doc)
    if extension in {".docx", ".doc"}:
        try:
            content = extract_docx_content(file_path)
            return {
                "success": True,
                "type": "text",
                "filename": filename,
                "file_type": "docx",
                "page_number": 1,
                "total_pages": 1,
                "is_multipage": False,
                "content": INVOICE_PROMPT_BANNER + content,
            }
        except Exception as error:
            try:
                FAILED_DIR.mkdir(parents=True, exist_ok=True)
                shutil.move(str(file_path), str(FAILED_DIR / filename))
                moved = True
            except Exception:
                moved = False
            return {
                "success": False,
                "error": f"Could not read Word document: {error}." + (" Moved to failed folder." if moved else ""),
                "moved_to_failed": moved,
            }

    # Excel spreadsheets (.xlsx, .xls)
    if extension in {".xlsx", ".xls"}:
        try:
            content = extract_excel_content(file_path)
            return {
                "success": True,
                "type": "text",
                "filename": filename,
                "file_type": "excel",
                "page_number": 1,
                "total_pages": 1,
                "is_multipage": False,
                "content": INVOICE_PROMPT_BANNER + content,
            }
        except Exception as error:
            try:
                FAILED_DIR.mkdir(parents=True, exist_ok=True)
                shutil.move(str(file_path), str(FAILED_DIR / filename))
                moved = True
            except Exception:
                moved = False
            return {
                "success": False,
                "error": f"Could not read Excel document: {error}." + (" Moved to failed folder." if moved else ""),
                "moved_to_failed": moved,
            }

    # PDF documents
    if extension == ".pdf":
        try:
            with fitz.open(file_path) as document:
                total_pages = len(document)
                if page_number < 1 or page_number > total_pages:
                    return {
                        "success": False,
                        "error": f"Invalid page_number {page_number}. Document has {total_pages} page(s).",
                    }
                page = document[page_number - 1]
                text = page.get_text().strip()
                has_embedded_images = len(page.get_images(full=True)) > 0
                if len(text) >= 50 and not has_embedded_images:
                    return {
                        "success": True,
                        "type": "text",
                        "filename": filename,
                        "page_number": page_number,
                        "total_pages": total_pages,
                        "is_multipage": total_pages > 1,
                        "content": INVOICE_PROMPT_BANNER + text,
                    }
                # Page has an embedded image (e.g. a scanned/photographed invoice pasted
                # into a "print to PDF" page) - always render it, even if the page's text
                # layer is long enough to pass the threshold above. Text-only extraction
                # would otherwise silently return incidental page text (timestamps,
                # filenames, URLs from the source webpage) instead of the real invoice.
                img = optimize_image_payload(page)
                banner_text = INVOICE_PROMPT_BANNER + f"[Document: {filename} | Page {page_number} of {total_pages}]"
                if text:
                    banner_text += (
                        "\n\n[Extracted text layer below - may be incidental page text "
                        "(timestamp/filename/URL) rather than the invoice itself; the "
                        "invoice content is in the attached image:]\n" + text
                    )
                return [
                    TextContent(type="text", text=banner_text),
                    img,
                ]
        except Exception as error:
            try:
                FAILED_DIR.mkdir(parents=True, exist_ok=True)
                shutil.move(str(file_path), str(FAILED_DIR / filename))
                moved = True
            except Exception:
                moved = False
            return {
                "success": False,
                "error": f"Could not read PDF document: {error}." + (" Moved to failed folder." if moved else ""),
                "moved_to_failed": moved,
            }

    try:
        img_payload = (
            ImageContent(data=base64.b64encode(file_path.read_bytes()).decode("ascii"), mimeType="image/jpeg")
            if file_path.stat().st_size <= 800 * 1024 and extension in {".jpg", ".jpeg"}
            else optimize_image_payload(file_path)
        )
        return [
            TextContent(type="text", text=INVOICE_PROMPT_BANNER + f"[Image Document: {filename}]"),
            img_payload,
        ]
    except Exception as error:
        try:
            FAILED_DIR.mkdir(parents=True, exist_ok=True)
            shutil.move(str(file_path), str(FAILED_DIR / filename))
            moved = True
        except Exception:
            moved = False
        return {
            "success": False,
            "error": f"Could not read image document: {error}." + (" Moved to failed folder." if moved else ""),
            "moved_to_failed": moved,
        }


@mcp.tool()
def mark_invoice_as_failed(filename: str, reason: str = "") -> dict:
    """Move an unprocessable, corrupted, or failed invoice file from inbox to the failed directory.

    Args:
        filename: Name of the invoice file in the inbox.
        reason: Explanation of why the invoice could not be processed (e.g. 'Unreadable scan', 'Corrupted', 'Not an invoice').
    """
    try:
        file_path = inbox_file(filename)
        FAILED_DIR.mkdir(parents=True, exist_ok=True)
        target_path = FAILED_DIR / filename
        shutil.move(str(file_path), str(target_path))

        payload = load_json_payload()
        record = {
            "filename": display_filename(filename),
            "status": "failed",
            "reason": reason,
            "source": "claude_mcp_client",
            "failed_at": datetime.now(timezone.utc).isoformat(),
            "file_moved_to_failed": True,
        }
        records = {make_record_key(item): item for item in payload.get("invoices", [])}
        records[f"{filename}|failed"] = record
        payload["generated_at"] = datetime.now(timezone.utc).isoformat()
        payload["invoices"] = list(records.values())
        save_json_payload(payload)

        return {
            "success": True,
            "message": f"Moved {filename} to failed folder. Reason: {reason}",
            "failed_path": str(target_path),
        }
    except Exception as error:
        return {"success": False, "error": str(error)}


def make_record_key(item: dict) -> str:
    """Generate a unique key per invoice entry even within the same multi-page PDF."""
    fname = item.get("filename", "")
    page = item.get("page_number", 1)
    inv_no = item.get("REF_NO") or item.get("invoice no.(Ai)", "")
    return f"{fname}|page_{page}|inv_{inv_no}"


CANONICAL_SERVICES = {
    "Rent_997212": "RENT-2026",
    "Maintenance.": "Contract-2026",
    "AMC Charges.": "Contract-2026",
    "PARKING RENT": "Contract-2026",
    "Rent-(Urd)": "Rent-(Urd)",
    "Commission Charges-Store_997221": "COMMISSION",
    "Electricity Chrg": "Contract-2026",
    "General Expenses": "Contract-2026",
}

CANONICAL_TDS = {
    "RENT-2026",
    "Contract-2026",
    "Rent-(Urd)",
    "COMMISSION",
}


def normalize_tds_and_service(tds_name: str = "", service_name: str = "", hint_text: str = "") -> tuple[str, str]:
    """Normalize TDS_NAME and SERVICE_NAME to canonical Sheet1 values based on input and hints."""
    t = (tds_name or "").strip()
    s = (service_name or "").strip()

    # Direct match if already exact canonical service
    for canon_srv, canon_tds in CANONICAL_SERVICES.items():
        if s.lower() == canon_srv.lower():
            final_tds = t if t in CANONICAL_TDS else canon_tds
            return final_tds, canon_srv

    h = (hint_text or "").strip().lower()
    combo = f"{t} {s} {h}".lower()

    if "parking" in combo:
        return "Contract-2026", "PARKING RENT"
    if "amc" in combo or "hvac" in combo or "lift" in combo:
        return "Contract-2026", "AMC Charges."
    if "cam" in combo or "maintenance" in combo:
        return "Contract-2026", "Maintenance."
    if "commission" in combo:
        return "COMMISSION", "Commission Charges-Store_997221"
    if "electricity" in combo or "power" in combo or " dg " in f" {combo} ":
        return "Contract-2026", "Electricity Chrg"
    if "urd" in combo or "unregistered" in combo:
        return "Rent-(Urd)", "Rent-(Urd)"
    if "general" in combo or "expense" in combo:
        return "Contract-2026", "General Expenses"
    if "rent" in combo or "lease" in combo or "licence" in combo or "license" in combo or "997212" in combo:
        return "RENT-2026", "Rent_997212"

    return t, s


def format_date_m_d_yyyy(value: object) -> str:
    """Format invoice date into M/D/YYYY (e.g. 9/1/2026, 9/2/2026, 8/1/2026) without leading zeros."""
    if not value:
        return ""
    if isinstance(value, (datetime, date)):
        return f"{value.month}/{value.day}/{value.year}"

    val_str = str(value).strip()
    if not val_str or val_str.upper() in {"N/A", "NA", "NONE", "NULL", "-"}:
        return val_str

    m_exact = re.fullmatch(r"([1-9]|1[0-2])/([1-9]|[12][0-9]|3[01])/(\d{4})", val_str)
    if m_exact:
        return val_str

    date_formats = [
        "%d-%m-%Y", "%d/%m/%Y", "%d.%m.%Y",
        "%d-%m-%y", "%d/%m/%y", "%d.%m.%y",
        "%d-%b-%Y", "%d-%b-%y", "%d-%B-%Y", "%d-%B-%y",
        "%d %b %Y", "%d %b %y", "%d %B %Y", "%d %B %y",
        "%Y-%m-%d", "%Y/%m/%d",
        "%Y-%m-%d %H:%M:%S",
    ]
    for fmt in date_formats:
        try:
            dt = datetime.strptime(val_str, fmt)
            return f"{dt.month}/{dt.day}/{dt.year}"
        except ValueError:
            continue

    m = re.search(r"(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})", val_str)
    if m:
        p1, p2, p3 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        year = p3 + 2000 if p3 < 100 else p3
        if p2 <= 12 and p1 <= 31:
            day, month = p1, p2
        elif p1 <= 12 and p2 <= 31:
            month, day = p1, p2
        else:
            month, day = min(p1, 12), min(p2, 31)
        try:
            dt = datetime(year, month, day)
            return f"{dt.month}/{dt.day}/{dt.year}"
        except ValueError:
            pass

    return val_str


def lookup_sheet1_vendor_rules(gst_number: str = "", pan_number: str = "") -> list[dict]:
    """Load vendor matching rules from Sheet1 of invoice_ledger.xlsx for the given GST/PAN."""
    excel_path = INVOICES_DIR / "invoice_ledger.xlsx"
    if not excel_path.exists():
        return []
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True, read_only=True)
        if "Sheet1" not in wb.sheetnames:
            wb.close()
            return []
        sheet = wb["Sheet1"]
        rows_iter = sheet.iter_rows(values_only=True)
        header_row = None
        for r in rows_iter:
            if any(str(c).strip().lower() == "gst number" for c in r if c):
                header_row = [str(c).strip().lower() if c else "" for c in r]
                break
        if not header_row:
            wb.close()
            return []
        gst_idx = header_row.index("gst number") if "gst number" in header_row else -1
        pan_idx = header_row.index("pan number") if "pan number" in header_row else -1
        tds_idx = header_row.index("tds_name") if "tds_name" in header_row else -1
        srv_idx = header_row.index("service_name") if "service_name" in header_row else -1
        amt_idx = header_row.index("amount") if "amount" in header_row else -1
        tag_amt_idx = header_row.index("tag_admsite_amount") if "tag_admsite_amount" in header_row else -1
        paybale_idx = header_row.index("paybale") if "paybale" in header_row else -1

        results = []
        for r in rows_iter:
            r_gst = str(r[gst_idx]).strip().upper() if gst_idx >= 0 and gst_idx < len(r) and r[gst_idx] else ""
            r_pan = str(r[pan_idx]).strip().upper() if pan_idx >= 0 and pan_idx < len(r) and r[pan_idx] else ""
            if (gst_number and r_gst == gst_number) or (pan_number and r_pan == pan_number):
                results.append({
                    "tds_name": r[tds_idx] if tds_idx >= 0 and tds_idx < len(r) else "",
                    "service_name": r[srv_idx] if srv_idx >= 0 and srv_idx < len(r) else "",
                    "amount": r[amt_idx] if amt_idx >= 0 and amt_idx < len(r) else None,
                    "tag_admsite_amount": r[tag_amt_idx] if tag_amt_idx >= 0 and tag_amt_idx < len(r) else None,
                    "paybale": r[paybale_idx] if paybale_idx >= 0 and paybale_idx < len(r) else None,
                })
        wb.close()
        return results
    except Exception:
        return []


INVALID_VENDOR_PREFIX = re.compile(
    r"^(?:v[\s\-]*mart|v-mart|retail\b|rent\b|rental\b|lease\b|\d+|\b(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|jun|jul|aug|sep|sept|oct|nov|dec)\b)",
    re.I,
)


def clean_vendor_name(n: str) -> str:
    if not n:
        return ""
    n = re.sub(
        r"^(?:duplicate copy of [^.]*\.?\s*|duplicate of [^.]*\.?\s*|duplicate copy\.?\s*|page \d+ is blank[^.]*\.?\s*|unregistered vendor\s*\(no gst\)\s*,\s*|blank trailing page[^.]*\.?\s*)",
        "",
        n,
        flags=re.I,
    ).strip()
    n = re.sub(r"^(?:(?:co-)?landlords?|vendors?|suppliers?|billers?|from)\s*[:\-]?\s*", "", n, flags=re.I).strip()
    n = re.sub(r"\s*\(landlord\)", "", n, flags=re.I).strip()
    n = re.sub(r"\s*\(unregistered[^)]*\)", "", n, flags=re.I).strip()
    if "(" in n and ")" not in n:
        n = n.split("(")[0].strip()
    return n.strip(" ,.-")


def is_valid_vendor_name(name: str) -> bool:
    if not name or len(name) < 3:
        return False
    if re.search(r"\bv[\s\-]*mart\b", name, re.I):
        return False
    if INVALID_VENDOR_PREFIX.match(name):
        return False
    if re.match(
        r"^(?:total|subtotal|basic|amount|charges|bill|taxable|invoice|period|store|premises|property|room|hall|shop|floor|flat|unit|page|date|renting|lease|rental|common|maintenance|electricity|amenities|ac|hvac|licence|license)\b",
        name,
        re.I,
    ):
        return False
    return True


def extract_vendor_from_remark(rem: str) -> str:
    if not rem:
        return ""

    DELIM = r"(?:,|\;|\bstore\b|\bpremises\b|\bproperty\b|\bhs\b|\bpan\b|\bgstin\b|\bbuyer\b|\btenant\b|\bfor\s+v[\s\-]*mart|(?<!\b[A-Z])\.(?:\s+|$)|$)"

    # 1. '<Name> - <Description>' at beginning
    m_dash = re.match(r"^([A-Za-z0-9\.\,\&\s\(\)\'\/\-]+?)\s*[-–]\s*(?:Rent|Electricity|Maintenance|Lease|CAM|Rental|Bill|Common|Charges|Tufanganj|Shri|Padrauna|Aligarh|Building|House|Ward)", rem, re.I)
    if m_dash:
        c = clean_vendor_name(m_dash.group(1))
        if is_valid_vendor_name(c):
            return c

    # 2. Explicit keyword: Vendor: / Landlord: / Supplier: / Co-landlord: / Biller:
    m = re.search(rf"(?:vendor|landlord|supplier|co-landlord|biller|landlords)\s*(?:name)?\s*[:\-]\s*([A-Za-z0-9\.\,\&\s\(\)\'\/\-]+?){DELIM}", rem, re.I)
    if m:
        c = clean_vendor_name(m.group(1))
        if is_valid_vendor_name(c):
            return c

    # 3. 'co-owner Landlord \d+ of \d+: <Name>'
    m_co = re.search(rf"co-owner\s+landlord\s+\d+\s+of\s+\d+\s*:\s*([A-Za-z0-9\.\,\&\s\(\)\'\/\-]+?){DELIM}", rem, re.I)
    if m_co:
        c = clean_vendor_name(m_co.group(1))
        if is_valid_vendor_name(c):
            return c

    # 4. '<Name>\'s individual co-owner rent invoice'
    m_ind = re.search(r"([A-Za-z0-9\.\,\&\s\(\)\'\/\-]+?)\'s\s+(?:individual\s+)?co-owner", rem, re.I)
    if m_ind:
        c = clean_vendor_name(m_ind.group(1))
        if is_valid_vendor_name(c):
            return c

    # 5. 'landlord <Name>' or 'vendor <Name>'
    m = re.search(rf"\b(?:landlord|vendor|supplier|co-landlord)\s+([A-Z][A-Za-z0-9\.\,\&\s\(\)\'\/\-]+?){DELIM}", rem, re.I)
    if m:
        c = clean_vendor_name(m.group(1))
        if is_valid_vendor_name(c):
            return c

    # 6. 'bill from <Name>' or 'invoice from <Name>' or 'invoices from <Name>'
    m_from = re.search(rf"(?:bill|bills|invoice|invoices|rent)\s+from\s+([A-Za-z0-9\.\,\&\s\(\)\'\/\-]+?){DELIM}", rem, re.I)
    if m_from:
        c = clean_vendor_name(m_from.group(1))
        if is_valid_vendor_name(c):
            return c

    # 7. '... to <date>[\),]\s*<Name>,'
    m_dt = re.search(r"(?:to\s+[0-9]{1,2}[-\.\/][0-9]{1,2}[-\.\/][0-9]{2,4}|to\s+[0-9]{1,2}[-\.\s][a-z]{3,9}[-\.\s][0-9]{2,4})[\)\],\s]+\s*([^,;]+)", rem, re.I)
    if m_dt:
        c = clean_vendor_name(m_dt.group(1))
        if is_valid_vendor_name(c):
            return c

    # 8. 'Rent / Rental / CAM ... for <period> - <Name>'
    m_period = re.search(r"(?:(?:to|by|period|month of)\s+[0-9\.\-\/a-z\s]+?|(?:rent|rental|charges|income|bill|services)\s+(?:receivable\s+)?for\s+(?:the\s+month\s+of\s+)?(?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z0-9\s\.\-\/\']*)[,\-–]\s*([^,;]+)", rem, re.I)
    if m_period:
        c = clean_vendor_name(m_period.group(1))
        if is_valid_vendor_name(c):
            return c

    # 9. 'Rent on Immovable Property, <Name>' or 'Rent, <Name>'
    m_rent_imm = re.search(r"(?:rent\s*(?:on\s+immovable\s+property|paid|of\s+premises)?|common\s+area\s+maintenance\s+charges)\s*[,]\s*([^,;]+)", rem, re.I)
    if m_rent_imm:
        c = clean_vendor_name(m_rent_imm.group(1))
        if is_valid_vendor_name(c):
            return c

    # 10. 'for <month> <year>, <Name>'
    m_mth = re.search(r"for\s+(?:january|february|march|april|may|june|july|august|september|october|november|december|jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)\s+\d{2,4}\s*[,]\s*([^,;]+)", rem, re.I)
    if m_mth:
        c = clean_vendor_name(m_mth.group(1))
        if is_valid_vendor_name(c):
            return c

    # 11. 'CAM Charges - <Name>'
    m_cam = re.search(r"(?:cam\s+charges|maintenance\s+charges)\s*[-–]\s*([^,;]+)", rem, re.I)
    if m_cam:
        c = clean_vendor_name(m_cam.group(1))
        if is_valid_vendor_name(c):
            return c

    # 12. Mall name reference: '<Name> Mall <City>'
    m_mall = re.search(r"\b([A-Z][A-Za-z0-9\s]+?Mall\s+[A-Z][A-Za-z]+)", rem)
    if m_mall:
        c = clean_vendor_name(m_mall.group(1))
        if is_valid_vendor_name(c):
            return c

    # 13. First segment before comma/period
    first_seg = re.split(r"[,.\;]", rem)[0].strip()
    first_seg = clean_vendor_name(first_seg)
    if 3 < len(first_seg) < 60 and is_valid_vendor_name(first_seg):
        return first_seg

    return ""


@mcp.tool()
def save_extracted_data_to_json(
    filename: str,
    invoice_number_ai: str,
    invoice_date_ai: str,
    amount_ai: float | None,
    tds_name: str,
    service_name: str,
    vendor_name: str = "",
    gst_number: str = "",
    pan_number: str = "",
    remark: str = "",
    page_number: int = 1,
    total_pages: int = 1,
) -> dict:
    """Save Claude's extraction to JSON, then move the source inbox file to processed/ (never deleted) when all pages are done.

    IMPORTANT FOR CLAUDE AI:
    Analyze the invoice document (line items, description, header, SAC codes, tax tables) to extract:
    - vendor_name (REQUIRED):
        * Vendor, Seller, Supplier, or Landlord name extracted DIRECTLY from the invoice header/document.
        * Do NOT use Sheet1; extract the name directly from the invoice scan.
    - amount_ai (REQUIRED):
        * Must be the TOTAL TAXABLE VALUE (Taxable Amount / Base Amount before GST taxes).
        * Look for 'Total Taxable Value', 'Taxable Amount', 'Taxable Value', 'Basic Amount', or 'Subtotal'.
        * Do NOT pass the Grand Total (inclusive of GST). The Amount column in the ledger holds the Taxable Value.
    - tds_name (REQUIRED):
        * 'RENT-2026' for premises rent, lease, licence fees
        * 'Contract-2026' for CAM, maintenance, AMC, parking, electricity, or contractor services
        * 'Rent-(Urd)' for unregistered vendor rent
        * 'COMMISSION' for store commission
    - service_name (REQUIRED):
        * 'Rent_997212' for premises rent (SAC 997212)
        * 'Maintenance.' for Common Area Maintenance (CAM) / maintenance charges
        * 'AMC Charges.' for Annual Maintenance Contracts (HVAC, Lift, DG, equipment)
        * 'PARKING RENT' for parking space rental
        * 'Electricity Chrg' for electricity / power backup charges
        * 'Commission Charges-Store_997221' for store commission
        * 'Rent-(Urd)' for unregistered vendor rental
        * 'General Expenses' for general contractor expenses

    Args:
        filename: Name of the invoice file in the inbox.
        invoice_number_ai: Invoice / bill number extracted from the document.
        invoice_date_ai: Invoice date extracted from the document.
        amount_ai: REQUIRED Total Taxable Value (Taxable Amount / Base Amount before GST taxes).
        tds_name: REQUIRED TDS category decided by analyzing this invoice.
        service_name: REQUIRED Service category decided by analyzing this invoice.
        gst_number: 15-character GST number (GSTIN) if present. If missing, leave empty.
        pan_number: 10-character PAN number. If GST number is missing, provide PAN number so vendor can be matched via PAN.
        remark: Any remarks, notes, description, billing period, or payment terms extracted from the invoice file.
        page_number: The page number this invoice was extracted from (default: 1).
        total_pages: Total number of pages in the PDF document (default: 1).
    """
    try:
        file_path = inbox_file(filename)
        payload = load_json_payload()

        # If PDF and total_pages was not explicitly specified (or is 1), auto-detect actual page count
        if file_path.suffix.lower() == ".pdf" and total_pages <= 1:
            try:
                with fitz.open(file_path) as doc:
                    if len(doc) > 1:
                        total_pages = len(doc)
            except Exception:
                pass

        is_multipage = total_pages > 1

        gst_cleaned = gst_number.strip().upper() if gst_number else ""
        pan_cleaned = pan_number.strip().upper() if pan_number else ""

        # If PAN is empty but GST is provided, auto-derive PAN from GSTIN if possible (chars 3 to 12)
        if not pan_cleaned and gst_cleaned:
            pan_match = re.search(r"[A-Z]{5}[0-9]{4}[A-Z]", gst_cleaned)
            if pan_match:
                pan_cleaned = pan_match.group(0)

        # Normalize TDS_NAME and SERVICE_NAME based on Claude's input and hints
        hint = f"{filename} {invoice_number_ai} {remark}"
        clean_tds, clean_service = normalize_tds_and_service(tds_name, service_name, hint)

        # Check Sheet1 rules if multiple categories exist for this vendor (e.g. Rent vs Maintenance with same GST)
        sheet1_candidates = lookup_sheet1_vendor_rules(gst_cleaned, pan_cleaned)
        if sheet1_candidates:
            unique_srvs = {
                (str(c.get("tds_name") or "").strip().lower(), str(c.get("service_name") or "").strip().lower())
                for c in sheet1_candidates
                if c.get("service_name") or c.get("tds_name")
            }
            if len(unique_srvs) > 1:
                best_cand = None
                if amount_ai is not None:
                    try:
                        amt_val = float(amount_ai)
                        for cand in sheet1_candidates:
                            for af in ("amount", "tag_admsite_amount", "paybale"):
                                cv = cand.get(af)
                                if cv not in (None, ""):
                                    try:
                                        c_num = float(str(cv).replace(",", "").replace("₹", "").strip())
                                        if c_num > 0 and abs(amt_val - c_num) < 2.0:
                                            best_cand = cand
                                            break
                                    except (ValueError, TypeError):
                                        pass
                            if best_cand:
                                break
                    except (ValueError, TypeError):
                        pass

                if not best_cand:
                    h_lower = hint.lower()
                    for cand in sheet1_candidates:
                        c_srv = str(cand.get("service_name") or "").strip().lower()
                        if any(k in h_lower for k in ("maintenance", "cam", "amenities")) and "maintenance" in c_srv:
                            best_cand = cand
                            break
                        if any(k in h_lower for k in ("amc", "hvac", "lift")) and "amc" in c_srv:
                            best_cand = cand
                            break
                        if "parking" in h_lower and "parking" in c_srv:
                            best_cand = cand
                            break
                        if any(k in h_lower for k in ("electricity", "power")) and "electricity" in c_srv:
                            best_cand = cand
                            break
                        if any(k in h_lower for k in ("rent", "lease", "licence")) and "rent" in c_srv:
                            best_cand = cand
                            break

                if best_cand:
                    clean_tds = str(best_cand.get("tds_name") or clean_tds).strip()
                    clean_service = str(best_cand.get("service_name") or clean_service).strip()

        # If Claude somehow passed empty values, infer from hints and filename
        if not clean_tds or not clean_service:
            inf_tds, inf_srv = normalize_tds_and_service("", "", hint)
            clean_tds = clean_tds or inf_tds or "RENT-2026"
            clean_service = clean_service or inf_srv or "Rent_997212"

        clean_date = format_date_m_d_yyyy(invoice_date_ai)

        v_name = clean_vendor_name(vendor_name)
        if not v_name:
            v_name = extract_vendor_from_remark(remark)
        if not v_name:
            stem = re.sub(r"[\(\)\d_\-]+", " ", Path(filename).stem).strip()
            stem = re.sub(r"\b(v[\s\-]*mart|pdf|jpg|jpeg|png|scan|doc|wa|bill|invoice|sep|sept|september|aug|august|rent|rental|cam|inv|new|copy|retail|ltd|limited)\b", "", stem, flags=re.I).strip()
            if len(stem) > 3:
                v_name = stem.title()

        record = {
            "filename": display_filename(filename),
            "status": "saved",
            "source": "claude_mcp_client",
            "NAME": v_name,
            "vendor_name": v_name,
            "REF_NO": invoice_number_ai.strip(),
            "REF_DT": clean_date,
            "invoice no.(Ai)": invoice_number_ai.strip(),
            "invoice date(AI)": clean_date,
            "AMOUNT(AI)": amount_ai,
            "GST NUMBER": gst_cleaned,
            "PAN NUMBER": pan_cleaned,
            "TDS_NAME": clean_tds,
            "SERVICE_NAME": clean_service,
            "Remark": remark.strip(),
            "Remarks": remark.strip(),
            "page_number": page_number,
            "total_pages": total_pages,
            "is_multipage": is_multipage,
            "saved_at": datetime.now(timezone.utc).isoformat(),
        }
        records = {make_record_key(item): item for item in payload.get("invoices", [])}
        records[make_record_key(record)] = record
        payload["generated_at"] = datetime.now(timezone.utc).isoformat()
        payload["invoices"] = list(records.values())
        save_json_payload(payload)

        # Only move the source file out of inbox once we've reached the last page
        # (or single-page document). It's moved to processed/, never deleted, so the
        # original scan/photo stays available for reference.
        source_moved = False
        if page_number >= total_pages:
            PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
            target_path = PROCESSED_DIR / file_path.name
            counter = 1
            while target_path.exists():
                target_path = PROCESSED_DIR / f"{file_path.stem} ({counter}){file_path.suffix}"
                counter += 1
            shutil.move(str(file_path), str(target_path))
            source_moved = True
            record["file_moved_to_processed"] = True
            record["processed_path"] = str(target_path)
            record["processed_at"] = datetime.now(timezone.utc).isoformat()
            payload["generated_at"] = datetime.now(timezone.utc).isoformat()
            save_json_payload(payload)

        return {
            "success": True,
            "json_file": str(EXTRACTED_DATA_PATH),
            "page_number": page_number,
            "total_pages": total_pages,
            "source_file_moved_to_processed": source_moved,
            "message": (
                f"Saved invoice for {filename} (page {page_number}/{total_pages}). "
                + ("Source file moved to processed/." if source_moved else "More pages remain in this PDF.")
            ),
        }
    except Exception as error:
        return {"success": False, "error": str(error)}


@mcp.tool()
def get_extracted_invoices() -> dict:
    """Read the saved extracted_invoices.json data without changing it."""
    try:
        return {"success": True, **load_json_payload()}
    except Exception as error:
        return {"success": False, "error": str(error)}


if __name__ == "__main__":
    mcp.run()
