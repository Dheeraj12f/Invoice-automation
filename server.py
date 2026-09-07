"""MCP server for extracting invoice data into a local JSON file.

Claude reads invoice scans through MCP. This server stores the values Claude
extracts; it does not call an external AI API and does not use Excel.
"""

from __future__ import annotations

import base64
import json
import re
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone
from pathlib import Path

import fitz
import openpyxl
from mcp.server.mcpserver import MCPServer
from mcp.types import ImageContent, TextContent

mcp = MCPServer("Invoice Processing Server")

BASE_DIR = Path(__file__).resolve().parent
INVOICES_DIR = BASE_DIR / "invoices"
INBOX_DIR = INVOICES_DIR / "inbox"
PROCESSED_DIR = INVOICES_DIR / "processed"  # Retained for compatibility; not used.
FAILED_DIR = INVOICES_DIR / "failed"
EXTRACTED_DATA_PATH = BASE_DIR / "extracted_invoices.json"

for directory in (INBOX_DIR, PROCESSED_DIR, FAILED_DIR):
    directory.mkdir(parents=True, exist_ok=True)


def inbox_file(filename: str) -> Path:
    """Return a direct file in inbox and reject path traversal."""
    path = (INBOX_DIR / filename).resolve()
    if path.parent != INBOX_DIR.resolve() or not path.is_file():
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
    supported = {".pdf", ".png", ".jpg", ".jpeg", ".docx", ".doc", ".xlsx", ".xls"}
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
                "content": content,
            }
        except Exception as error:
            return {"success": False, "error": str(error)}

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
                "content": content,
            }
        except Exception as error:
            return {"success": False, "error": str(error)}

    # PDF documents
    if extension == ".pdf":
        with fitz.open(file_path) as document:
            total_pages = len(document)
            if page_number < 1 or page_number > total_pages:
                return {
                    "success": False,
                    "error": f"Invalid page_number {page_number}. Document has {total_pages} page(s).",
                }
            page = document[page_number - 1]
            text = page.get_text().strip()
            if len(text) >= 50:
                return {
                    "success": True,
                    "type": "text",
                    "filename": filename,
                    "page_number": page_number,
                    "total_pages": total_pages,
                    "is_multipage": total_pages > 1,
                    "content": text,
                }
            img = optimize_image_payload(page)
            if total_pages > 1:
                return [
                    TextContent(
                        type="text",
                        text=(
                            f"[Document: {filename} | Page {page_number} of {total_pages}. "
                            f"Use read_invoice_content(filename='{filename}', page_number=...) to read other pages.]"
                        ),
                    ),
                    img,
                ]
            return img

    if file_path.stat().st_size <= 800 * 1024 and extension in {".jpg", ".jpeg"}:
        return ImageContent(
            data=base64.b64encode(file_path.read_bytes()).decode("ascii"),
            mimeType="image/jpeg",
        )
    return optimize_image_payload(file_path)


def make_record_key(item: dict) -> str:
    """Generate a unique key per invoice entry even within the same multi-page PDF."""
    fname = item.get("filename", "")
    page = item.get("page_number", 1)
    inv_no = item.get("invoice no.(Ai)", "")
    return f"{fname}|page_{page}|inv_{inv_no}"


@mcp.tool()
def save_extracted_data_to_json(
    filename: str,
    invoice_number_ai: str,
    invoice_date_ai: str,
    amount_ai: float | None,
    gst_number: str = "",
    pan_number: str = "",
    page_number: int = 1,
    total_pages: int = 1,
) -> dict:
    """Save Claude's extraction to JSON, then delete the source inbox file when all pages are done.

    Args:
        filename: Name of the invoice file in the inbox.
        invoice_number_ai: Invoice / bill number extracted from the document.
        invoice_date_ai: Invoice date extracted from the document.
        amount_ai: Total invoice amount.
        gst_number: 15-character GST number (GSTIN) if present. If missing, leave empty.
        pan_number: 10-character PAN number. If GST number is missing, provide the PAN number so the vendor details can be matched via PAN.
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

        record = {
            "filename": filename,
            "status": "saved",
            "source": "claude_mcp_client",
            "invoice no.(Ai)": invoice_number_ai.strip(),
            "invoice date(AI)": invoice_date_ai.strip(),
            "AMOUNT(AI)": amount_ai,
            "GST NUMBER": gst_cleaned,
            "PAN NUMBER": pan_cleaned,
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

        # Only delete source file if we have reached the last page (or single-page document)
        source_deleted = False
        if page_number >= total_pages:
            file_path.unlink()
            source_deleted = True
            record["file_deleted"] = True
            record["deleted_at"] = datetime.now(timezone.utc).isoformat()
            payload["generated_at"] = datetime.now(timezone.utc).isoformat()
            save_json_payload(payload)

        return {
            "success": True,
            "json_file": str(EXTRACTED_DATA_PATH),
            "page_number": page_number,
            "total_pages": total_pages,
            "source_file_deleted": source_deleted,
            "message": (
                f"Saved invoice for {filename} (page {page_number}/{total_pages}). "
                + ("Source file deleted." if source_deleted else "More pages remain in this PDF.")
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
