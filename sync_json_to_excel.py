"""Continuously copy newly saved invoice JSON records into the Excel ledger."""

from __future__ import annotations

import json
import os
import re
import shutil
import time
from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

BASE_DIR = Path(__file__).resolve().parent
JSON_PATH = BASE_DIR / "extracted_invoices.json"
INVOICES_DIR = BASE_DIR / "invoices"
EXCEL_PATH = INVOICES_DIR / "invoice_ledger.xlsx"
INBOX_DIR = INVOICES_DIR / "inbox"
FAILED_DIR = INVOICES_DIR / "failed"
POLL_SECONDS = 2
SHEET1_HEADER_ROW = 2

HEADERS = [
    "SUPPLIER_SLID", "NAME", "REF_ADMSITE_SHRTNAME", "TDS_NAME", "REF_NO", "REF_DT",
    "SERVICE_NAME", "AMOUNT", "AMOUNT(SHEET1)", "TAG_ADMSITE_SHRTNAME", "TAG_ADMSITE_AMOUNT",
    "TDS", "GST", "Paybale", "TERM", "FORM_NAME",
    "Remark", "GST NO.", "PAN NUMBER",
]

COL_INDEX = {name: idx for idx, name in enumerate(HEADERS)}
IDX_SUPPLIER_SLID = COL_INDEX["SUPPLIER_SLID"]
IDX_NAME = COL_INDEX["NAME"]
IDX_REF_ADMSITE_SHRTNAME = COL_INDEX["REF_ADMSITE_SHRTNAME"]
IDX_TDS_NAME = COL_INDEX["TDS_NAME"]
IDX_REF_NO = COL_INDEX["REF_NO"]
IDX_REF_DT = COL_INDEX["REF_DT"]
IDX_SERVICE_NAME = COL_INDEX["SERVICE_NAME"]
IDX_AMOUNT = COL_INDEX["AMOUNT"]
IDX_AMOUNT_SHEET1 = COL_INDEX["AMOUNT(SHEET1)"]
IDX_TAG_ADMSITE_SHRTNAME = COL_INDEX["TAG_ADMSITE_SHRTNAME"]
IDX_TAG_ADMSITE_AMOUNT = COL_INDEX["TAG_ADMSITE_AMOUNT"]
IDX_TDS = COL_INDEX["TDS"]
IDX_GST = COL_INDEX["GST"]
IDX_PAYBALE = COL_INDEX["Paybale"]
IDX_TERM = COL_INDEX["TERM"]
IDX_FORM_NAME = COL_INDEX["FORM_NAME"]
IDX_REMARK = COL_INDEX["Remark"]
IDX_REMARKS = IDX_REMARK
IDX_GST_NO = COL_INDEX["GST NO."]
IDX_PAN_NO = COL_INDEX["PAN NUMBER"]

GREEN_HEADER_FILL = PatternFill("solid", fgColor="A9D08E")
FLAGGED_FILL = PatternFill("solid", fgColor="FFD9D9")
FLAGGED_FONT = Font(color="9C0006")
HEADER_FONT_BLACK = Font(bold=True, color="000000")
HEADER_BORDER = Border(
    left=Side(style="thin", color="000000"),
    right=Side(style="thin", color="000000"),
    top=Side(style="thin", color="000000"),
    bottom=Side(style="thin", color="000000"),
)


def normalise_pan_number(value: object) -> str:
    """Compare PAN numbers consistently (10 alphanumeric uppercase characters)."""
    if not value:
        return ""
    val = str(value).upper()
    match = re.search(r"[A-Z]{5}[0-9]{4}[A-Z]", val)
    if match:
        return match.group(0)
    cleaned = "".join(character for character in val if character.isalnum())
    return cleaned if len(cleaned) == 10 else ""


def normalise_gst_number(value: object) -> str:
    """Compare GST numbers consistently despite case or formatting differences."""
    if not value:
        return ""
    val = str(value).upper()
    match = re.search(r"[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]", val)
    if match:
        return match.group(0)
    cleaned = "".join(character for character in val if character.isalnum())
    return cleaned if len(cleaned) == 15 else ""


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


# Statutory rates observed consistently across Sheet1's existing matched rows, used only
# to estimate TDS/GST/Paybale for a vendor that has NO Sheet1 master entry at all (no GST
# or PAN match). GST is a flat 18% (CGST 9% + SGST 9%, or IGST 18%) across every category.
# TDS varies by category: RENT-2026 rows are consistently 10%; Contract-2026 rows split
# between 1% and 2% in Sheet1 (194C: 1% for individual/HUF contractors, 2% for others) -
# 2% is used as the more common default; COMMISSION is 2%; Rent-(Urd) is 0% (unregistered
# rent, no TDS deducted per existing Sheet1 rows).
GST_RATE = 0.18

TDS_RATE_BY_TDS_NAME = {
    "rent-2026": 0.10,
    "contract-2026": 0.02,
    "commission": 0.02,
    "rent-(urd)": 0.0,
}

# GST state codes (first 2 digits of a GSTIN) mapped to state/UT name, used to derive a
# 'VRL- <State>' admin-site short name when there is no Sheet1 master row to supply one.
GST_STATE_CODES = {
    "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab", "04": "Chandigarh",
    "05": "Uttarakhand", "06": "Haryana", "07": "Delhi", "08": "Rajasthan", "09": "Uttar Pradesh",
    "10": "Bihar", "11": "Sikkim", "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur",
    "15": "Mizoram", "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
    "20": "Jharkhand", "21": "Odisha", "22": "Chhattisgarh", "23": "Madhya Pradesh", "24": "Gujarat",
    "26": "Dadra and Nagar Haveli and Daman and Diu", "27": "Maharashtra",
    "29": "Karnataka", "30": "Goa", "31": "Lakshadweep", "32": "Kerala", "33": "Tamil Nadu",
    "34": "Puducherry", "35": "Andaman and Nicobar Islands", "36": "Telangana", "37": "Andhra Pradesh",
    "38": "Ladakh",
}

STATE_NAME_TOKENS = {
    name.lower(): name for name in GST_STATE_CODES.values()
} | {
    "up": "Uttar Pradesh", "mp": "Madhya Pradesh", "ap": "Andhra Pradesh", "tn": "Tamil Nadu",
    "wb": "West Bengal", "hp": "Himachal Pradesh", "j&k": "Jammu and Kashmir",
}


def round_amount(val: float | int | str | None) -> float | None:
    """Round off amount to nearest integer (e.g. 221197.88 -> 221198.00, 23.4 -> 23.00)."""
    if val is None or val == "":
        return None
    try:
        num = float(str(val).replace(",", "").replace("₹", "").strip())
        d = Decimal(str(num)).quantize(Decimal("1"), rounding=ROUND_HALF_UP)
        return float(d)
    except Exception:
        return val


def estimate_tds_gst_paybale(amount: float | int | None, tds_name: str) -> tuple[float, float, float]:
    """Best-effort TDS/GST/Paybale for a vendor with no Sheet1 master row, using the
    statutory rate table above. Returns ("", "", "") if amount is missing/invalid."""
    try:
        amt = float(amount)
    except (TypeError, ValueError):
        return "", "", ""
    if amt <= 0:
        return "", "", ""
    rate = TDS_RATE_BY_TDS_NAME.get((tds_name or "").strip().lower(), 0.0)
    tds = round(amt * rate, 2)
    gst = round(amt * GST_RATE, 2)
    paybale = round(amt + gst - tds, 2)
    return tds, gst, paybale


def derive_admsite_shrtname(gst_number: str = "", pan_number: str = "") -> str:
    """Best-effort state-level admin-site short name (e.g. 'VRL- Uttar Pradesh') derived
    from the GST state code, used only when there is no Sheet1 master row to supply it."""
    gst = normalise_gst_number(gst_number)
    state_code = gst[:2] if gst else ""
    state_name = GST_STATE_CODES.get(state_code)
    return f"VRL- {state_name}" if state_name else ""


def extract_site_from_remark(remark: str) -> str:
    """Best-effort store/site short name pulled from the invoice remark Claude already
    wrote after reading the document (e.g. '...Digiha Tiraha, Gonda Road, Bahraich, UP.'
    -> 'BAHRAICH'), used only when there is no Sheet1 master row to supply
    TAG_ADMSITE_SHRTNAME. Looks for the segment immediately before a recognized Indian
    state name/abbreviation."""
    if not remark:
        return ""
    parts = [p.strip() for p in re.split(r"[,.]", remark) if p.strip()]
    for i, part in enumerate(parts):
        if part.lower() in STATE_NAME_TOKENS and i > 0:
            candidate = parts[i - 1]
            if 2 < len(candidate) <= 40 and not re.search(r"\d", candidate):
                return candidate.upper()

    m = re.search(r"(?:store|premises|property)\s*[:\-]?\s*([A-Za-z][A-Za-z\s]{2,30}?)(?:,|\.|$)", remark, re.I)
    if m:
        return m.group(1).strip().upper()

    return ""


def find_sheet1_header_row(sheet, max_scan: int = 10) -> int:
    """Locate the row holding Sheet1's column headers by scanning for 'GST NUMBER' /
    'PAN NUMBER' text, instead of trusting the hardcoded SHEET1_HEADER_ROW. A manual
    edit in Excel (e.g. deleting a blank spacer row) can shift the real header up or
    down, and a stale row number silently breaks every Sheet1 match."""
    for row_idx in range(1, max_scan + 1):
        values = [str(c.value).strip().lower() if c.value is not None else "" for c in sheet[row_idx]]
        if "gst number" in values or "pan number" in values:
            return row_idx
    return SHEET1_HEADER_ROW


def load_sheet1_lookup(
    workbook,
) -> tuple[dict[str, list[dict[str, object]]], dict[str, list[dict[str, object]]], dict[str, list[dict[str, object]]]]:
    """Return Sheet1 rows indexed by GST NUMBER, PAN NUMBER, and vendor NAME values."""
    if "Sheet1" not in workbook.sheetnames:
        raise RuntimeError(
            "Sheet1 is missing. Add the data-validation tab named Sheet1 "
            "before syncing invoice records."
        )

    sheet = workbook["Sheet1"]
    header_row = find_sheet1_header_row(sheet)
    source_headers = [
        str(cell.value).strip() if cell.value is not None else ""
        for cell in sheet[header_row]
    ]
    header_positions = {
        header.lower(): index for index, header in enumerate(source_headers) if header
    }
    gst_position = header_positions.get("gst number")
    pan_position = header_positions.get("pan number")
    name_position = header_positions.get("name")
    if gst_position is None and pan_position is None:
        raise RuntimeError(
            "Sheet1 must have a 'GST NUMBER' or 'PAN NUMBER' header in one of its first 10 rows."
        )

    gst_lookup: dict[str, list[dict[str, object]]] = {}
    pan_lookup: dict[str, list[dict[str, object]]] = {}
    name_lookup: dict[str, list[dict[str, object]]] = {}
    for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        row_dict = {
            header.lower(): values[index] if index < len(values) else ""
            for index, header in enumerate(source_headers) if header
        }
        if gst_position is not None and gst_position < len(values):
            gst_number = normalise_gst_number(values[gst_position])
            if gst_number:
                gst_lookup.setdefault(gst_number, []).append(row_dict)
        if pan_position is not None and pan_position < len(values):
            pan_number = normalise_pan_number(values[pan_position])
            if pan_number:
                pan_lookup.setdefault(pan_number, []).append(row_dict)
        if name_position is not None and name_position < len(values):
            name_val = str(values[name_position]).strip().lower() if values[name_position] else ""
            if name_val:
                name_lookup.setdefault(name_val, []).append(row_dict)

    return gst_lookup, pan_lookup, name_lookup


def select_best_sheet1_row(
    candidates: list[dict[str, object]],
    target_service: str = "",
    target_tds: str = "",
    invoice_amount: float | int | None = None,
    hint_text: str = "",
    used_candidates: list[dict[str, object]] | None = None,
) -> dict[str, object] | None:
    """Select the best matching Sheet1 row for a vendor.

    If Sheet1 has multiple rows with different TDS_NAME or SERVICE_NAME for the same GST:
    1. Match both target_service and target_tds (canonical master pairing).
    2. Check for exact/close amount match against Sheet1 amount, tag_admsite_amount, or paybale.
    3. Check for keyword matches in hint_text (maintenance, cam, amc, parking, electricity, rent).
    4. Match target_service / target_tds.
    5. Avoid re-using candidate rows already claimed by another invoice for this same GST.
    """
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    # Filter to unused candidates if some candidates were already used
    available = candidates
    if used_candidates:
        unused = [c for c in candidates if c not in used_candidates]
        if unused:
            available = unused

    # 0. Match by both service_name and tds_name (canonical combination in Sheet1)
    if target_service and target_tds:
        t_srv = target_service.strip().lower()
        t_tds = target_tds.strip().lower()
        exact_both = [
            c for c in available
            if str(c.get("service_name", "")).strip().lower() == t_srv
            and str(c.get("tds_name", "")).strip().lower() == t_tds
        ]
        if len(exact_both) == 1:
            return exact_both[0]
        if len(exact_both) > 1:
            available = exact_both

    # 1. Direct Amount Matching (highest precision for differentiating Rent vs Maintenance)
    if invoice_amount is not None:
        try:
            inv_amt_val = float(str(invoice_amount).replace(",", "").replace("₹", "").strip())
            if inv_amt_val > 0:
                best_amt_cand = None
                smallest_diff = float("inf")
                for cand in available:
                    for amt_field in ("amount", "tag_admsite_amount", "paybale"):
                        val = cand.get(amt_field)
                        if val not in (None, ""):
                            try:
                                c_val = float(str(val).replace(",", "").replace("₹", "").strip())
                                if c_val > 0:
                                    diff = abs(inv_amt_val - c_val)
                                    if diff < 2.0:  # Direct exact or rounding match
                                        return cand
                                    if diff < smallest_diff and (diff / c_val) < 0.02:  # within 2%
                                        smallest_diff = diff
                                        best_amt_cand = cand
                            except (ValueError, TypeError):
                                pass
                if best_amt_cand is not None:
                    return best_amt_cand
        except (ValueError, TypeError):
            pass

    # 2. Keyword Matching in hints (filename, remark, invoice description)
    h = (hint_text or "").lower()
    if h:
        for cand in available:
            c_srv = str(cand.get("service_name", "")).strip().lower()
            c_tds = str(cand.get("tds_name", "")).strip().lower()
            if any(k in h for k in ("maintenance", "cam", "amenities", "common area")) and (
                "maintenance" in c_srv or "maintenance" in c_tds
            ):
                return cand
            if any(k in h for k in ("amc", "hvac", "lift", "dg amc")) and "amc" in c_srv:
                return cand
            if "parking" in h and "parking" in c_srv:
                return cand
            if any(k in h for k in ("electricity", "power", "dg unit")) and "electricity" in c_srv:
                return cand
            if any(k in h for k in ("commission", "brokerage")) and "commission" in c_srv:
                return cand
            if any(k in h for k in ("rent", "lease", "licence", "license", "997212")) and (
                "rent" in c_srv or "rent" in c_tds
            ):
                return cand

    # 3. Match by exact canonical service_name
    if target_service:
        t_srv = target_service.strip().lower()
        for cand in available:
            if str(cand.get("service_name", "")).strip().lower() == t_srv:
                return cand

    # 4. Match by exact canonical tds_name
    if target_tds:
        t_tds = target_tds.strip().lower()
        for cand in available:
            if str(cand.get("tds_name", "")).strip().lower() == t_tds:
                return cand

    return available[0]


def get_last_data_row(sheet) -> int:
    """Return the 1-based index of the last row containing real non-empty data."""
    for row_idx in range(sheet.max_row, 1, -1):
        if any(sheet.cell(row=row_idx, column=c).value not in (None, "") for c in range(1, len(HEADERS) + 1)):
            return row_idx
    return 1


def get_missing_ai_fields(row: list) -> list[str]:
    """Check if any AI extracted field (REF_NO, REF_DT, AMOUNT) has a missing, unreadable, or invalid value."""
    missing = []
    # Check REF_NO (column index IDX_REF_NO)
    inv_no = str(row[IDX_REF_NO]).strip() if row[IDX_REF_NO] is not None else ""
    if not inv_no or inv_no.upper() in {"N/A", "NA", "NONE", "NOT FOUND", "NOT LEGIBLE", "OCR UNCLEAR", "NULL", "-"}:
        missing.append("REF_NO")

    # Check REF_DT (column index IDX_REF_DT)
    inv_date = str(row[IDX_REF_DT]).strip() if row[IDX_REF_DT] is not None else ""
    if not inv_date or inv_date.upper() in {"N/A", "NA", "NONE", "NOT FOUND", "NOT LEGIBLE", "NULL", "-"}:
        missing.append("REF_DT")

    # Check AMOUNT (column index IDX_AMOUNT)
    amt = row[IDX_AMOUNT]
    if amt is None or str(amt).strip() in {"", "None", "null", "N/A", "NA", "NONE"}:
        missing.append("AMOUNT")
    else:
        try:
            val_num = float(str(amt).replace(",", "").replace("₹", "").strip())
            if val_num <= 0:
                missing.append("AMOUNT")
        except ValueError:
            missing.append("AMOUNT")

    return missing


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

    # 1. '<Name> - <Description>' at beginning of remark
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


def extract_vendor_name_from_invoice(record: dict) -> str:
    """Extract vendor name directly from invoice fields (NAME, vendor_name, Remark, or filename).
    Per user requirement, the vendor name must come from the invoice ONLY, NEVER from Sheet1."""
    name = str(record.get("NAME") or record.get("vendor_name") or record.get("supplier_name") or "").strip()
    if name and is_valid_vendor_name(name):
        return name

    remark = str(record.get("Remark") or record.get("Remarks") or record.get("remark") or "").strip()
    extracted = extract_vendor_from_remark(remark)
    if extracted:
        return extracted

    fname = Path(str(record.get("filename") or "")).stem
    stem = re.sub(r"[\(\)\d_\-]+", " ", fname).strip()
    stem = re.sub(r"\b(v[\s\-]*mart|pdf|jpg|jpeg|png|scan|doc|wa|bill|invoice|sep|sept|september|aug|august|rent|rental|cam|inv|new|copy|retail|ltd|limited)\b", "", stem, flags=re.I).strip()
    if len(stem) > 3:
        return stem.title()

    return ""


extract_vendor_name_fallback = extract_vendor_name_from_invoice


def compact_invoice_sheet(sheet) -> None:
    """Eliminate empty gap rows so data rows start immediately at row 2 with zero gaps."""
    if sheet.max_row <= 1:
        return

    # Extract all rows that contain at least one non-empty cell value
    data_rows = []
    for r in range(2, sheet.max_row + 1):
        vals = [sheet.cell(row=r, column=c).value for c in range(1, len(HEADERS) + 1)]
        if any(v not in (None, "") for v in vals):
            is_red = False
            cell_fill = sheet.cell(row=r, column=1).fill
            if cell_fill and cell_fill.fill_type == "solid":
                fill_color = str(getattr(cell_fill.fgColor, "rgb", ""))
                if "FFD9D9" in fill_color or "FFC7CE" in fill_color:
                    is_red = True

            missing_ai = get_missing_ai_fields(vals)
            if missing_ai:
                is_red = True

            data_rows.append((vals, is_red))

    # Check if there is a gap or trailing empty rows (i.e. if max_row != 1 + len(data_rows))
    last_row = get_last_data_row(sheet)
    has_gaps = (
        (sheet.max_row != 1 + len(data_rows))
        or (last_row != 1 + len(data_rows))
        or any(
            all(sheet.cell(row=r, column=c).value in (None, "") for c in range(1, len(HEADERS) + 1))
            for r in range(2, last_row + 1)
        )
    )

    if not has_gaps:
        return

    # Clear everything from row 2 downwards
    sheet.delete_rows(2, sheet.max_row)

    # Re-write the data rows compactly starting at row 2
    for idx, (vals, is_red) in enumerate(data_rows, start=2):
        cur_tds = str(vals[IDX_TDS_NAME] or "")
        cur_service = str(vals[IDX_SERVICE_NAME] or "")
        cur_hint = f"{vals[IDX_NAME]} {vals[IDX_REF_NO]} {vals[IDX_REMARK]}"
        norm_tds, norm_service = normalize_tds_and_service(cur_tds, cur_service, cur_hint)
        vals[IDX_TDS_NAME] = norm_tds
        vals[IDX_SERVICE_NAME] = norm_service
        vals[IDX_REF_DT] = format_date_m_d_yyyy(vals[IDX_REF_DT])

        for c, val in enumerate(vals, start=1):
            cell = sheet.cell(row=idx, column=c, value=val)
            if is_red:
                cell.fill = FLAGGED_FILL
                cell.font = FLAGGED_FONT
        sheet.cell(row=idx, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'

    sheet.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{1 + len(data_rows)}"


def reconcile_existing_invoices(sheet, gst_lookup: dict, pan_lookup: dict, name_lookup: dict | None = None) -> int:
    """Inspect all rows in Invoices sheet to:
    1. Correct any duplicate or mismatched TDS_NAME, SERVICE_NAME, AMOUNT(SHEET1), and amounts when Sheet1 has different entries for the same GST.
    2. Deduplicate identical rows (same REF_NO and same GST).
    """
    if sheet.max_row <= 1:
        return 0

    rows_data = []
    for r in range(2, sheet.max_row + 1):
        vals = [sheet.cell(row=r, column=c).value for c in range(1, len(HEADERS) + 1)]
        if any(v not in (None, "") for v in vals):
            rows_data.append(vals)

    if not rows_data:
        return 0

    # 1. Deduplicate identical REF_NO + GST rows (e.g. DDIVMA47 original & duplicate pages)
    seen_invoices = set()
    deduped_rows = []
    removed_duplicates = 0
    for vals in rows_data:
        ref_no = str(vals[IDX_REF_NO] or "").strip().lower()
        gst_no = str(vals[IDX_GST_NO] or "").strip().upper()
        amt = str(vals[IDX_AMOUNT] or "").strip()
        key = (ref_no, gst_no, amt)
        if ref_no and gst_no and key in seen_invoices:
            removed_duplicates += 1
            continue
        if ref_no and gst_no:
            seen_invoices.add(key)
        deduped_rows.append(vals)

    updated_count = 0
    # 2. Ensure all REF_DT values are formatted to M/D/YYYY (e.g. 9/1/2026)
    for vals in deduped_rows:
        cur_dt = vals[IDX_REF_DT]
        formatted_dt = format_date_m_d_yyyy(cur_dt)
        if formatted_dt and str(cur_dt).strip() != formatted_dt:
            vals[IDX_REF_DT] = formatted_dt
            updated_count += 1

    # 3. For rows sharing the same GST, verify and correct Sheet1 matching
    by_gst: dict[str, list[tuple[int, list]]] = defaultdict(list)
    for idx, vals in enumerate(deduped_rows):
        gst = normalise_gst_number(vals[IDX_GST_NO])
        pan = normalise_pan_number(vals[IDX_PAN_NO]) or normalise_pan_number(vals[IDX_GST_NO])
        if not pan and gst and len(gst) >= 12:
            pan = gst[2:12]
        k = gst or pan
        if k:
            by_gst[k].append((idx, vals))
        elif vals[IDX_NAME]:
            by_gst[str(vals[IDX_NAME]).strip().lower()].append((idx, vals))

    for key_id, row_items in by_gst.items():
        candidates = gst_lookup.get(key_id) or pan_lookup.get(key_id)
        if not candidates and len(key_id) >= 12:
            candidates = pan_lookup.get(key_id[2:12])
        if not candidates and name_lookup:
            candidates = name_lookup.get(key_id.lower())
            if not candidates and row_items:
                c_name = str(row_items[0][1][IDX_NAME] or "").strip().lower()
                candidates = name_lookup.get(c_name)
        if not candidates:
            # No Sheet1 master row exists for this GST/PAN. Backfill the admin-site and
            # TDS/GST/Paybale columns from the invoice's own data instead of leaving
            # them blank, and note in Remark that no Sheet1 match exists.
            for row_idx, vals in row_items:
                changed = False
                if vals[IDX_AMOUNT] not in (None, ""):
                    rnd_amt = round_amount(vals[IDX_AMOUNT])
                    if rnd_amt is not None and rnd_amt != vals[IDX_AMOUNT]:
                        vals[IDX_AMOUNT] = rnd_amt
                        changed = True
                new_ref_site = derive_admsite_shrtname(str(vals[IDX_GST_NO] or ""))
                if new_ref_site and vals[IDX_REF_ADMSITE_SHRTNAME] != new_ref_site:
                    vals[IDX_REF_ADMSITE_SHRTNAME] = new_ref_site
                    changed = True
                new_tag_site = extract_site_from_remark(str(vals[IDX_REMARK] or ""))
                if new_tag_site and vals[IDX_TAG_ADMSITE_SHRTNAME] != new_tag_site:
                    vals[IDX_TAG_ADMSITE_SHRTNAME] = new_tag_site
                    changed = True
                est_tds, est_gst, est_paybale = estimate_tds_gst_paybale(vals[IDX_AMOUNT], str(vals[IDX_TDS_NAME] or ""))
                if est_tds != "":
                    if vals[IDX_TAG_ADMSITE_AMOUNT] != vals[IDX_AMOUNT]:
                        vals[IDX_TAG_ADMSITE_AMOUNT] = vals[IDX_AMOUNT]
                        changed = True
                    if vals[IDX_TDS] != est_tds:
                        vals[IDX_TDS] = est_tds
                        changed = True
                    if vals[IDX_GST] != est_gst:
                        vals[IDX_GST] = est_gst
                        changed = True
                    if vals[IDX_PAYBALE] != est_paybale:
                        vals[IDX_PAYBALE] = est_paybale
                        changed = True
                rem_str = str(vals[IDX_REMARK] or "")
                if "GST No. not available in Sheet1" not in rem_str and "NO MATCH IN SHEET1" not in rem_str:
                    vals[IDX_REMARK] = (
                        f"{rem_str} | GST No. not available in Sheet1 (master directory) - "
                        "TAG_ADMSITE/TDS/GST estimated from invoice"
                    ).strip(" |")
                    changed = True
                if changed:
                    updated_count += 1
            continue

        used_candidates = []
        for row_idx, vals in row_items:
            inv_amt = vals[IDX_AMOUNT]
            inv_ref = str(vals[IDX_REF_NO] or "")
            inv_rem = str(vals[IDX_REMARK] or "")
            hint = f"{vals[IDX_NAME]} {inv_ref} {inv_rem}"

            matched_cand = select_best_sheet1_row(
                candidates,
                target_service=str(vals[IDX_SERVICE_NAME] or ""),
                target_tds=str(vals[IDX_TDS_NAME] or ""),
                invoice_amount=inv_amt,
                hint_text=hint,
                used_candidates=used_candidates if len(candidates) > 1 else None,
            )
            if matched_cand:
                if len(candidates) > 1:
                    used_candidates.append(matched_cand)

                changed = False

                if vals[IDX_AMOUNT] not in (None, ""):
                    rnd_amt = round_amount(vals[IDX_AMOUNT])
                    if rnd_amt is not None and rnd_amt != vals[IDX_AMOUNT]:
                        vals[IDX_AMOUNT] = rnd_amt
                        changed = True

                # Populate / synchronize SUPPLIER_SLID
                new_slid = str(matched_cand.get("supplier_slid") or "").strip()
                if new_slid and str(vals[IDX_SUPPLIER_SLID] or "").strip() != new_slid:
                    vals[IDX_SUPPLIER_SLID] = new_slid
                    changed = True

                # Populate / synchronize AMOUNT(SHEET1)
                new_s1_amt = matched_cand.get("amount")
                if new_s1_amt not in (None, "") and vals[IDX_AMOUNT_SHEET1] != new_s1_amt:
                    vals[IDX_AMOUNT_SHEET1] = new_s1_amt
                    changed = True

                new_tds = str(matched_cand.get("tds_name") or "").strip()
                new_srv = str(matched_cand.get("service_name") or "").strip()
                new_site = matched_cand.get("tag_admsite_shrtname") or ""
                new_tag_amt = matched_cand.get("tag_admsite_amount")
                if new_tag_amt in (None, ""):
                    new_tag_amt = matched_cand.get("amount") or ""
                new_tds_amt = matched_cand.get("tds") or ""
                new_gst_amt = matched_cand.get("gst") or ""
                new_paybale = matched_cand.get("paybale") or ""
                new_term = matched_cand.get("term") or ""
                new_form = matched_cand.get("form_name") or ""

                if str(vals[IDX_TDS_NAME] or "").strip() != new_tds:
                    vals[IDX_TDS_NAME] = new_tds
                    changed = True
                if str(vals[IDX_SERVICE_NAME] or "").strip() != new_srv:
                    vals[IDX_SERVICE_NAME] = new_srv
                    changed = True
                if new_tag_amt not in (None, "") and vals[IDX_TAG_ADMSITE_AMOUNT] != new_tag_amt:
                    vals[IDX_TAG_ADMSITE_AMOUNT] = new_tag_amt
                    changed = True
                if new_tds_amt not in (None, "") and vals[IDX_TDS] != new_tds_amt:
                    vals[IDX_TDS] = new_tds_amt
                    changed = True
                if new_gst_amt not in (None, "") and vals[IDX_GST] != new_gst_amt:
                    vals[IDX_GST] = new_gst_amt
                    changed = True
                if new_paybale not in (None, "") and vals[IDX_PAYBALE] != new_paybale:
                    vals[IDX_PAYBALE] = new_paybale
                    changed = True
                if new_site and vals[IDX_TAG_ADMSITE_SHRTNAME] != new_site:
                    vals[IDX_TAG_ADMSITE_SHRTNAME] = new_site
                    changed = True
                if new_term and vals[IDX_TERM] != new_term:
                    vals[IDX_TERM] = new_term
                    changed = True
                if new_form and vals[IDX_FORM_NAME] != new_form:
                    vals[IDX_FORM_NAME] = new_form
                    changed = True

                if changed:
                    updated_count += 1

    if removed_duplicates > 0 or updated_count > 0:
        sheet.delete_rows(2, sheet.max_row)
        for idx, vals in enumerate(deduped_rows, start=2):
            missing_ai = get_missing_ai_fields(vals)
            rem_text = str(vals[IDX_REMARK] or "")
            is_red = (
                len(missing_ai) > 0
                or "No matching record in Sheet1" in rem_text
                or "not available in Sheet1" in rem_text
                or "NO MATCH IN SHEET1" in rem_text
                or not vals[IDX_GST_NO]
            )
            for c, val in enumerate(vals, start=1):
                cell = sheet.cell(row=idx, column=c, value=val)
                if is_red:
                    cell.fill = FLAGGED_FILL
                    cell.font = FLAGGED_FONT
            sheet.cell(row=idx, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=idx, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=idx, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=idx, column=IDX_TDS + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=idx, column=IDX_GST + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=idx, column=IDX_PAYBALE + 1).number_format = '"₹"#,##0.00'
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{1 + len(deduped_rows)}"

    # Reconcile existing duplicate sheet rows to round off AMOUNT
    if "DUPLICATE DATA" in sheet.parent.sheetnames:
        dup_s = sheet.parent["DUPLICATE DATA"]
        for r in range(2, dup_s.max_row + 1):
            cur_amt = dup_s.cell(row=r, column=IDX_AMOUNT + 1).value
            if cur_amt not in (None, ""):
                rnd_amt = round_amount(cur_amt)
                if rnd_amt is not None and rnd_amt != cur_amt:
                    dup_s.cell(row=r, column=IDX_AMOUNT + 1, value=rnd_amt)
                    updated_count += 1

    if updated_count > 0 or removed_duplicates > 0:
        print(f"Reconciled ledger: corrected {updated_count} row(s), removed {removed_duplicates} duplicate row(s).")

    return updated_count + removed_duplicates


def atomic_save_workbook(workbook, target_path: Path) -> None:
    """Save workbook atomically to prevent partial writes and file corruption."""
    target_path.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = target_path.with_suffix(".tmp_save")
    workbook.save(tmp_path)
    workbook.close()

    bak_path = target_path.with_suffix(".bak")
    if target_path.exists():
        try:
            shutil.copy2(target_path, bak_path)
        except Exception:
            pass
        try:
            os.replace(tmp_path, target_path)
        except PermissionError:
            if tmp_path.exists():
                tmp_path.unlink()
            raise PermissionError(f"{target_path.name} is currently open in Excel. Close it so the watcher can save.")
    else:
        os.replace(tmp_path, target_path)


def load_ledger_safely(excel_path: Path) -> Workbook:
    """Load workbook with auto-recovery from backup if file was corrupted or interrupted."""
    if not excel_path.exists():
        return Workbook()
    try:
        return load_workbook(excel_path)
    except Exception as error:
        bak_path = excel_path.with_suffix(".bak")
        if bak_path.exists():
            print(f"Notice: {excel_path.name} was corrupted ({error}). Recovering from backup {bak_path.name}...")
            shutil.copy2(bak_path, excel_path)
            return load_workbook(excel_path)
        raise RuntimeError(
            f"Could not open {excel_path.name}: {error}. File was corrupted during an interrupted write."
        )


def apply_header_styles(sheet) -> None:
    """Apply styling, borders, green accents, and column widths to header row 1."""
    for column in range(1, len(HEADERS) + 1):
        cell = sheet.cell(row=1, column=column)
        cell.font = HEADER_FONT_BLACK
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = HEADER_BORDER
        header_name = HEADERS[column - 1]
        if header_name in ("TDS", "GST", "Paybale"):
            cell.fill = GREEN_HEADER_FILL
        else:
            cell.fill = PatternFill(fill_type=None)

        if header_name in ("Remark", "Remarks"):
            sheet.column_dimensions[cell.column_letter].width = 30
        elif header_name in ("SUPPLIER_SLID", "GST NO.", "PAN NUMBER", "NAME", "SERVICE_NAME", "TAG_ADMSITE_SHRTNAME", "AMOUNT(SHEET1)"):
            sheet.column_dimensions[cell.column_letter].width = 22
        else:
            sheet.column_dimensions[cell.column_letter].width = 18


def migrate_legacy_sheet(sheet, old_headers: list) -> None:
    """Migrate in-place from any legacy columns to the 17 new columns."""
    old_col_map = {str(h).strip().lower(): i for i, h in enumerate(old_headers) if h}

    # Load Sheet1 lookup and JSON lookup for enrichment of Remarks, GST NO., and AMOUNT(SHEET1)
    s1_lookup = {}
    gst_lookup_full = {}
    pan_lookup_full = {}
    name_lookup_full = {}
    try:
        if "Sheet1" in sheet.parent.sheetnames:
            gst_lookup_full, pan_lookup_full, name_lookup_full = load_sheet1_lookup(sheet.parent)
            s1_sheet = sheet.parent["Sheet1"]
            s1_header_row = find_sheet1_header_row(s1_sheet)
            s1_headers = [str(c.value).strip().lower() for c in s1_sheet[s1_header_row] if c.value]
            s1_name_idx = s1_headers.index("name") if "name" in s1_headers else -1
            s1_gst_idx = s1_headers.index("gst number") if "gst number" in s1_headers else -1
            s1_slid_idx = s1_headers.index("supplier_slid") if "supplier_slid" in s1_headers else -1
            s1_rem_idx = s1_headers.index("remarks") if "remarks" in s1_headers else -1
            if s1_name_idx >= 0:
                for s1_row in s1_sheet.iter_rows(min_row=s1_header_row + 1, values_only=True):
                    if s1_name_idx < len(s1_row) and s1_row[s1_name_idx]:
                        nm = str(s1_row[s1_name_idx]).strip().lower()
                        gst_val = str(s1_row[s1_gst_idx]).strip() if (s1_gst_idx >= 0 and s1_gst_idx < len(s1_row) and s1_row[s1_gst_idx]) else ""
                        slid_val = str(s1_row[s1_slid_idx]).strip() if (s1_slid_idx >= 0 and s1_slid_idx < len(s1_row) and s1_row[s1_slid_idx]) else ""
                        rem_val = str(s1_row[s1_rem_idx]).strip() if (s1_rem_idx >= 0 and s1_rem_idx < len(s1_row) and s1_row[s1_rem_idx]) else ""
                        if nm not in s1_lookup or (gst_val and not s1_lookup[nm].get("gst")):
                            s1_lookup[nm] = {"gst": gst_val, "remarks": rem_val, "slid": slid_val}
                        elif slid_val and not s1_lookup[nm].get("slid"):
                            s1_lookup[nm]["slid"] = slid_val
    except Exception:
        pass

    json_by_ref = {}
    if JSON_PATH.exists():
        try:
            p_data = json.loads(JSON_PATH.read_text(encoding="utf-8"))
            for item in p_data.get("invoices", []):
                ref = item.get("REF_NO") or item.get("invoice no.(Ai)")
                if ref:
                    json_by_ref[str(ref).strip().lower()] = item
        except Exception:
            pass

    rows_data = []
    for r in range(2, sheet.max_row + 1):
        vals = [sheet.cell(row=r, column=c).value for c in range(1, len(old_headers) + 1)]
        if any(v not in (None, "") for v in vals):
            is_red = False
            cell_fill = sheet.cell(row=r, column=1).fill
            if cell_fill and cell_fill.fill_type == "solid":
                fill_color = str(getattr(cell_fill.fgColor, "rgb", ""))
                if "FFD9D9" in fill_color or "FFC7CE" in fill_color:
                    is_red = True

            new_vals = [""] * len(HEADERS)
            for new_idx, h_name in enumerate(HEADERS):
                h_lower = h_name.lower()
                if h_lower in old_col_map and old_col_map[h_lower] < len(vals):
                    new_vals[new_idx] = vals[old_col_map[h_lower]]

            # Map synonyms for GST NO.
            if not new_vals[IDX_GST_NO]:
                for alt in ("gst no.", "gst no", "gst number", "gstin"):
                    if alt in old_col_map and old_col_map[alt] < len(vals):
                        v = vals[old_col_map[alt]]
                        if v not in (None, "") and not isinstance(v, (int, float)):
                            new_vals[IDX_GST_NO] = v
                            break

            # Map synonyms for PAN NO. (from an old sheet that already had its own PAN column)
            if not new_vals[IDX_PAN_NO]:
                for alt in ("pan no.", "pan no", "pan number"):
                    if alt in old_col_map and old_col_map[alt] < len(vals):
                        v = vals[old_col_map[alt]]
                        if v not in (None, "") and not isinstance(v, (int, float)):
                            new_vals[IDX_PAN_NO] = v
                            break

            # A PAN was previously stuffed into GST NO. when the invoice had no GSTIN
            # (this column didn't exist yet). Move it to the dedicated PAN NO. column
            # instead of leaving it mislabeled as a GST number.
            gst_no_val = str(new_vals[IDX_GST_NO] or "").strip()
            if gst_no_val and not normalise_gst_number(gst_no_val):
                looks_like_pan = normalise_pan_number(gst_no_val)
                if looks_like_pan:
                    if not new_vals[IDX_PAN_NO]:
                        new_vals[IDX_PAN_NO] = looks_like_pan
                    new_vals[IDX_GST_NO] = ""

            # Map synonyms for Remark
            if not new_vals[IDX_REMARK]:
                for alt in ("remark", "remarks"):
                    if alt in old_col_map and old_col_map[alt] < len(vals):
                        v = vals[old_col_map[alt]]
                        if v not in (None, "") and str(v).strip().upper() != "GST":
                            new_vals[IDX_REMARK] = v
                            break

            # If old schema had AMOUNT(AI), prioritize it for AMOUNT
            amt_ai_idx = old_col_map.get("amount(ai)")
            if amt_ai_idx is not None and amt_ai_idx < len(vals) and vals[amt_ai_idx] not in (None, ""):
                new_vals[IDX_AMOUNT] = round_amount(vals[amt_ai_idx])
            elif new_vals[IDX_AMOUNT] not in (None, ""):
                new_vals[IDX_AMOUNT] = round_amount(new_vals[IDX_AMOUNT])

            # If old schema had AMOUNT(SHEET1), map it
            amt_s1_idx = old_col_map.get("amount(sheet1)")
            if amt_s1_idx is not None and amt_s1_idx < len(vals) and vals[amt_s1_idx] not in (None, ""):
                new_vals[IDX_AMOUNT_SHEET1] = vals[amt_s1_idx]

            # If old schema had invoice no.(Ai), prioritize it for REF_NO
            inv_no_idx = old_col_map.get("invoice no.(ai)")
            if inv_no_idx is not None and inv_no_idx < len(vals) and not new_vals[IDX_REF_NO]:
                new_vals[IDX_REF_NO] = vals[inv_no_idx]

            # If old schema had invoice date(AI), prioritize it for REF_DT
            inv_dt_idx = old_col_map.get("invoice date(ai)")
            if inv_dt_idx is not None and inv_dt_idx < len(vals) and not new_vals[IDX_REF_DT]:
                new_vals[IDX_REF_DT] = vals[inv_dt_idx]

            # Fallback enrichment for GST NO. and Remark from JSON and Sheet1
            ref_key = str(new_vals[IDX_REF_NO]).strip().lower() if new_vals[IDX_REF_NO] else ""
            nm_key = str(new_vals[IDX_NAME]).strip().lower() if new_vals[IDX_NAME] else ""
            inv_item = json_by_ref.get(ref_key, {})
            s1_info = s1_lookup.get(nm_key, {})

            if not new_vals[IDX_GST_NO]:
                if inv_item.get("GST NUMBER"):
                    new_vals[IDX_GST_NO] = inv_item.get("GST NUMBER")
                elif s1_info.get("gst"):
                    new_vals[IDX_GST_NO] = s1_info.get("gst")

            if not new_vals[IDX_PAN_NO] and inv_item.get("PAN NUMBER"):
                new_vals[IDX_PAN_NO] = inv_item.get("PAN NUMBER")

            # Clean dummy "GST" string from remark
            if str(new_vals[IDX_REMARK]).strip().upper() in ("GST", "NONE"):
                new_vals[IDX_REMARK] = ""

            rem_parts = []
            if new_vals[IDX_REMARK]:
                rem_parts.append(str(new_vals[IDX_REMARK]).strip())
            if inv_item:
                ai_rem = str(inv_item.get("Remark") or inv_item.get("Remarks") or inv_item.get("remark") or "").strip()
                if ai_rem and ai_rem not in rem_parts:
                    rem_parts.append(ai_rem)
                tot_pages = inv_item.get("total_pages", 1)
                pg_num = inv_item.get("page_number", 1)
                if inv_item.get("is_multipage") or tot_pages > 1:
                    rem_parts.append(f"Multi-page PDF (Page {pg_num} of {tot_pages})")
                if not inv_item.get("GST NUMBER"):
                    rem_parts.append("GST missing on invoice - matched via PAN")
            s1_rem = str(s1_info.get("remarks", "")).strip()
            if s1_rem and s1_rem.upper() not in ("NONE", "GST", "") and s1_rem not in rem_parts:
                rem_parts.append(s1_rem)
            if rem_parts:
                new_vals[IDX_REMARK] = " | ".join(rem_parts)

            # Ensure TDS_NAME and SERVICE_NAME are normalized and enriched from JSON if missing
            cur_tds = str(new_vals[IDX_TDS_NAME] or "")
            cur_service = str(new_vals[IDX_SERVICE_NAME] or "")
            if not cur_tds and inv_item.get("TDS_NAME"):
                cur_tds = str(inv_item.get("TDS_NAME"))
            if not cur_service and inv_item.get("SERVICE_NAME"):
                cur_service = str(inv_item.get("SERVICE_NAME"))
            cur_hint = f"{new_vals[IDX_NAME]} {new_vals[IDX_REF_NO]} {new_vals[IDX_REMARK]}"
            norm_tds, norm_service = normalize_tds_and_service(cur_tds, cur_service, cur_hint)
            new_vals[IDX_TDS_NAME] = norm_tds
            new_vals[IDX_SERVICE_NAME] = norm_service

            # Populate AMOUNT(SHEET1) based on GST/PAN, SERVICE_NAME, and TDS_NAME
            if not new_vals[IDX_AMOUNT_SHEET1]:
                m_gst = normalise_gst_number(new_vals[IDX_GST_NO])
                m_pan = normalise_pan_number(new_vals[IDX_PAN_NO]) or normalise_pan_number(new_vals[IDX_GST_NO])
                if not m_pan and m_gst and len(m_gst) >= 12:
                    m_pan = m_gst[2:12]
                cand_list = (gst_lookup_full.get(m_gst) if m_gst else None) or (pan_lookup_full.get(m_pan) if m_pan else None)
                if not cand_list and new_vals[IDX_NAME]:
                    cand_list = name_lookup_full.get(str(new_vals[IDX_NAME]).strip().lower())
                if cand_list:
                    m_cand = select_best_sheet1_row(
                        cand_list,
                        target_service=new_vals[IDX_SERVICE_NAME],
                        target_tds=new_vals[IDX_TDS_NAME],
                        invoice_amount=new_vals[IDX_AMOUNT],
                        hint_text=cur_hint,
                    )
                    if m_cand:
                        if m_cand.get("amount") not in (None, "") and not new_vals[IDX_AMOUNT_SHEET1]:
                            new_vals[IDX_AMOUNT_SHEET1] = m_cand.get("amount")
                        if m_cand.get("supplier_slid") and not new_vals[IDX_SUPPLIER_SLID]:
                            new_vals[IDX_SUPPLIER_SLID] = str(m_cand.get("supplier_slid") or "").strip()

            if not new_vals[IDX_SUPPLIER_SLID] and s1_info.get("slid"):
                new_vals[IDX_SUPPLIER_SLID] = str(s1_info.get("slid") or "").strip()

            missing_ai = get_missing_ai_fields(new_vals)
            if missing_ai:
                is_red = True
                missing_str = f"Missing in invoice: {', '.join(missing_ai)}"
                if not new_vals[IDX_REMARKS]:
                    new_vals[IDX_REMARKS] = missing_str
                elif "Missing in invoice:" not in str(new_vals[IDX_REMARKS]):
                    new_vals[IDX_REMARKS] = f"{new_vals[IDX_REMARKS]} | {missing_str}"

            rows_data.append((new_vals, is_red))

    sheet.delete_rows(1, sheet.max_row)
    sheet.append(HEADERS)
    apply_header_styles(sheet)
    for idx, (vals, is_red) in enumerate(rows_data, start=2):
        for c, val in enumerate(vals, start=1):
            cell = sheet.cell(row=idx, column=c, value=val)
            if is_red:
                cell.fill = FLAGGED_FILL
                cell.font = FLAGGED_FONT
        sheet.cell(row=idx, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_TDS + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_GST + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=idx, column=IDX_PAYBALE + 1).number_format = '"₹"#,##0.00'
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}{1 + len(rows_data)}"


def reconcile_and_move_existing_duplicates(sheet, dup_sheet) -> int:
    """Scan Invoices sheet for duplicate rows (same REF_NO, GST/Name, SERVICE_NAME, TDS_NAME),
    move duplicates to DUPLICATE DATA tab with normal styling (not red), and remove from Invoices."""
    if sheet.max_row <= 1:
        return 0

    seen: dict[tuple[str, str, str, str], int] = {}
    rows_to_remove = []
    moved_count = 0

    for r in range(2, sheet.max_row + 1):
        ref = str(sheet.cell(r, IDX_REF_NO + 1).value or "").strip().lower()
        gst = str(sheet.cell(r, IDX_GST_NO + 1).value or "").strip().upper()
        name = str(sheet.cell(r, IDX_NAME + 1).value or "").strip().lower()
        srv = str(sheet.cell(r, IDX_SERVICE_NAME + 1).value or "").strip().lower()
        tds = str(sheet.cell(r, IDX_TDS_NAME + 1).value or "").strip().lower()

        if not ref:
            continue

        vendor_key = gst if gst else name
        sig = (ref, vendor_key, srv, tds)

        if sig in seen:
            row_vals = [sheet.cell(r, c).value for c in range(1, len(HEADERS) + 1)]
            cur_rem = str(row_vals[IDX_REMARK] or "").strip()
            row_vals[IDX_REMARK] = f"{cur_rem} | DUPLICATE ROW (Matches Invoices row {seen[sig]})".strip(" |")

            target_row = get_last_data_row(dup_sheet) + 1
            for col_idx, val in enumerate(row_vals, start=1):
                cell = dup_sheet.cell(row=target_row, column=col_idx, value=val)
                cell.fill = PatternFill(fill_type=None)  # DON'T MARK RED
                cell.font = Font(color="000000")

            dup_sheet.cell(row=target_row, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
            dup_sheet.cell(row=target_row, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
            dup_sheet.cell(row=target_row, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'

            rows_to_remove.append(r)
            moved_count += 1
            print(f"Moved duplicate row {r} (Ref: {ref}, {vendor_key}) to 'DUPLICATE DATA' tab at row {target_row}.")
        else:
            seen[sig] = r

    for r in sorted(rows_to_remove, reverse=True):
        sheet.delete_rows(r, 1)

    return moved_count


def ensure_ledger():
    """Open the ledger, creating or migrating the expected invoice sheet and duplicate data sheet when needed."""
    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook = load_ledger_safely(EXCEL_PATH)
    sheet = workbook["Invoices"] if "Invoices" in workbook.sheetnames else workbook.active

    # Ensure DUPLICATE DATA sheet exists with proper headers and formatting
    if "DUPLICATE DATA" not in workbook.sheetnames:
        dup_sheet = workbook.create_sheet("DUPLICATE DATA")
    else:
        dup_sheet = workbook["DUPLICATE DATA"]

    dup_headers = [cell.value for cell in dup_sheet[1]] if dup_sheet.max_row else []
    if dup_headers != HEADERS:
        if dup_headers and any(dup_headers):
            migrate_legacy_sheet(dup_sheet, dup_headers)
        else:
            dup_sheet.delete_rows(1, max(dup_sheet.max_row, 1))
            dup_sheet.append(HEADERS)
            apply_header_styles(dup_sheet)
            dup_sheet.freeze_panes = "A2"
            dup_sheet.auto_filter.ref = f"A1:{get_column_letter(len(HEADERS))}1"
    else:
        compact_invoice_sheet(dup_sheet)

    headers = [cell.value for cell in sheet[1]] if sheet.max_row else []
    if headers != HEADERS:
        if headers and any(headers):
            migrate_legacy_sheet(sheet, headers)
        else:
            sheet.title = "Invoices"
            sheet.delete_rows(1, max(sheet.max_row, 1))
            sheet.append(HEADERS)
            apply_header_styles(sheet)
            sheet.freeze_panes = "A2"
            sheet.auto_filter.ref = sheet.dimensions
    else:
        # Compact any empty gap rows between header row 1 and data rows
        compact_invoice_sheet(sheet)

    if "SyncLog" not in workbook.sheetnames:
        log = workbook.create_sheet("SyncLog")
        log.append(["sync_key", "synced_at"])
        log.sheet_state = "hidden"
    return workbook, sheet, dup_sheet, workbook["SyncLog"]


def sync_once() -> int:
    """Write each unsynced saved JSON record to Excel, routing duplicates to DUPLICATE DATA tab."""
    if not JSON_PATH.exists():
        return 0
    try:
        payload = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return 0  # The server is still writing the JSON; retry on the next poll.

    invoices = payload.get("invoices", [])

    workbook, sheet, dup_sheet, log = ensure_ledger()
    gst_lookup, pan_lookup, name_lookup = load_sheet1_lookup(workbook)
    reconciled_count = reconcile_existing_invoices(sheet, gst_lookup, pan_lookup, name_lookup)
    moved_duplicates = reconcile_and_move_existing_duplicates(sheet, dup_sheet)
    logged_keys = {row[0] for row in log.iter_rows(min_row=2, values_only=True) if row[0]}
    synced_records = []

    # Map existing rows currently in Invoices sheet to avoid duplicate additions
    existing_invoices_in_sheet: set[tuple[str, str]] = set()
    for r in range(2, sheet.max_row + 1):
        ref_val = str(sheet.cell(r, IDX_REF_NO + 1).value or "").strip().lower()
        gst_val = str(sheet.cell(r, IDX_GST_NO + 1).value or "").strip().upper()
        name_val = str(sheet.cell(r, IDX_NAME + 1).value or "").strip().lower()
        if ref_val:
            if gst_val:
                existing_invoices_in_sheet.add((ref_val, gst_val))
            if name_val:
                existing_invoices_in_sheet.add((ref_val, name_val))

    # Map existing rows in DUPLICATE DATA sheet
    existing_in_dup_sheet: set[tuple[str, str]] = set()
    for r in range(2, dup_sheet.max_row + 1):
        ref_val = str(dup_sheet.cell(r, IDX_REF_NO + 1).value or "").strip().lower()
        gst_val = str(dup_sheet.cell(r, IDX_GST_NO + 1).value or "").strip().upper()
        name_val = str(dup_sheet.cell(r, IDX_NAME + 1).value or "").strip().lower()
        if ref_val:
            if gst_val:
                existing_in_dup_sheet.add((ref_val, gst_val))
            if name_val:
                existing_in_dup_sheet.add((ref_val, name_val))

    # Candidates: saved records not yet synced, OR missing from Invoices sheet
    candidates = []
    for record in invoices:
        if record.get("status") != "saved":
            continue
        # Skip trailing pages with 0 amount
        if record.get("page_number", 1) > 1 and float(record.get("AMOUNT(AI)", 0) or 0) == 0.0:
            continue
        inv_no_cand = str(record.get("REF_NO") or record.get("invoice no.(Ai)") or "").strip().lower()
        raw_gst_cand = record.get("GST NUMBER", "")
        clean_gst_cand = normalise_gst_number(raw_gst_cand).upper() if raw_gst_cand else ""
        v_name_cand = str(record.get("NAME") or extract_vendor_name_fallback(record) or "").strip().lower()

        in_invoices = False
        in_duplicates = False
        if inv_no_cand:
            if clean_gst_cand:
                in_invoices = (inv_no_cand, clean_gst_cand) in existing_invoices_in_sheet
                in_duplicates = (inv_no_cand, clean_gst_cand) in existing_in_dup_sheet
            elif v_name_cand:
                in_invoices = (inv_no_cand, v_name_cand) in existing_invoices_in_sheet
                in_duplicates = (inv_no_cand, v_name_cand) in existing_in_dup_sheet

        if not in_invoices or not in_duplicates or not record.get("excel_synced_at"):
            candidates.append(record)

    # 1. Keep existing rows in Invoices sheet synchronized with extracted_invoices.json (e.g. Taxable Value updates)
    json_by_ref_page: dict[tuple[str, int] | str, dict] = {}
    for rec in invoices:
        ref_k = str(rec.get("REF_NO") or rec.get("invoice no.(Ai)") or "").strip().lower()
        pg_k = rec.get("page_number", 1)
        json_by_ref_page[(ref_k, pg_k)] = rec
        if ref_k not in json_by_ref_page:
            json_by_ref_page[ref_k] = rec

    updated_existing = reconciled_count + moved_duplicates
    for r in range(2, sheet.max_row + 1):
        ref_val = str(sheet.cell(r, IDX_REF_NO + 1).value or "").strip()
        if not ref_val:
            continue
        rem_val = str(sheet.cell(r, IDX_REMARK + 1).value or "")
        pg_match = re.search(r"Page (\d+) of", rem_val)
        pg_num = int(pg_match.group(1)) if pg_match else 1

        rec = json_by_ref_page.get((ref_val.lower(), pg_num)) or json_by_ref_page.get(ref_val.lower())
        if rec:
            expected_amt = rec.get("AMOUNT(AI)")
            if expected_amt not in (None, ""):
                try:
                    exp_f = float(expected_amt)
                    cur_val = sheet.cell(r, IDX_AMOUNT + 1).value
                    cur_f = float(cur_val) if cur_val not in (None, "") else None
                    if cur_f is None or abs(cur_f - exp_f) > 0.001:
                        sheet.cell(r, IDX_AMOUNT + 1, value=exp_f)
                        sheet.cell(r, IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
                        updated_existing += 1
                except (ValueError, TypeError):
                    pass

    used_candidates_map: dict[str, list[dict]] = defaultdict(list)

    for record in candidates:
        page_num = record.get("page_number", 1)
        tot_pages = record.get("total_pages", 1)
        inv_no = record.get("REF_NO") or record.get("invoice no.(Ai)", "")
        inv_date = format_date_m_d_yyyy(record.get("REF_DT") or record.get("invoice date(AI)", ""))
        saved_at = record.get("saved_at", "")
        filename = record.get("filename", "")

        old_sync_key = f"{filename}|{saved_at}"
        new_sync_key = f"{filename}|page_{page_num}|inv_{inv_no}|{saved_at}"

        raw_gst = record.get("GST NUMBER", "")
        raw_pan = record.get("PAN NUMBER", "")

        clean_gst = normalise_gst_number(raw_gst)
        clean_pan = normalise_pan_number(raw_pan) or normalise_pan_number(raw_gst)

        # Determine AI decided TDS_NAME and SERVICE_NAME to pick the exact matching Sheet1 row
        ai_tds = str(record.get("TDS_NAME") or record.get("tds_name") or "").strip()
        ai_service = str(record.get("SERVICE_NAME") or record.get("service_name") or "").strip()
        hint = f"{filename} {inv_no} {record.get('Remark', '')} {record.get('Remarks', '')}"
        norm_ai_tds, norm_ai_service = normalize_tds_and_service(ai_tds, ai_service, hint)

        lookup_record = None
        matched_via = None
        cand_key = clean_gst or clean_pan
        if clean_gst and clean_gst in gst_lookup:
            lookup_record = select_best_sheet1_row(
                gst_lookup[clean_gst],
                target_service=norm_ai_service,
                target_tds=norm_ai_tds,
                invoice_amount=record.get("AMOUNT(AI)"),
                hint_text=hint,
                used_candidates=used_candidates_map[cand_key],
            )
            if lookup_record:
                used_candidates_map[cand_key].append(lookup_record)
            matched_via = f"GST {clean_gst}"
        elif clean_pan and clean_pan in pan_lookup:
            lookup_record = select_best_sheet1_row(
                pan_lookup[clean_pan],
                target_service=norm_ai_service,
                target_tds=norm_ai_tds,
                invoice_amount=record.get("AMOUNT(AI)"),
                hint_text=hint,
                used_candidates=used_candidates_map[cand_key],
            )
            if lookup_record:
                used_candidates_map[cand_key].append(lookup_record)
            matched_via = f"PAN {clean_pan}"
        elif record.get("NAME") and str(record.get("NAME")).strip().lower() in name_lookup:
            c_name = str(record.get("NAME")).strip().lower()
            lookup_record = select_best_sheet1_row(
                name_lookup[c_name],
                target_service=norm_ai_service,
                target_tds=norm_ai_tds,
                invoice_amount=record.get("AMOUNT(AI)"),
                hint_text=hint,
                used_candidates=used_candidates_map[cand_key],
            )
            if lookup_record:
                used_candidates_map[cand_key].append(lookup_record)
            matched_via = f"Name {record.get('NAME')}"

        # Check if this invoice is a duplicate of an existing row in Invoices tab
        # (Matches SERVICE NAME, TDS NAME, GST NO./Name, and REF_NO)
        existing_row_idx = None
        if inv_no:
            for r in range(2, sheet.max_row + 1):
                c_ref = str(sheet.cell(r, IDX_REF_NO + 1).value or "").strip().lower()
                c_gst = str(sheet.cell(r, IDX_GST_NO + 1).value or "").strip().upper()
                c_name = str(sheet.cell(r, IDX_NAME + 1).value or "").strip().lower()
                c_srv = str(sheet.cell(r, IDX_SERVICE_NAME + 1).value or "").strip().lower()
                c_tds = str(sheet.cell(r, IDX_TDS_NAME + 1).value or "").strip().lower()

                if c_ref == inv_no.lower():
                    is_same_vendor = False
                    if clean_gst and c_gst and c_gst == clean_gst.upper():
                        is_same_vendor = True
                    elif not clean_gst:
                        v_cand_name = str(record.get("NAME") or extract_vendor_name_fallback(record) or "").strip().lower()
                        if v_cand_name and (c_name == v_cand_name or v_cand_name in c_name):
                            is_same_vendor = True

                    if is_same_vendor:
                        if not c_srv or not norm_ai_service or c_srv == norm_ai_service.lower():
                            if not c_tds or not norm_ai_tds or c_tds == norm_ai_tds.lower():
                                existing_row_idx = r
                                break

        if existing_row_idx is not None:
            # User request: "IF THERE IS DUPLICATE ROWS IN INVOICES TAB LIKE SERVICE NAME AND TDS NAME, GST NO. AND OTHER COLUMNS .....THEN DON'T MARK RED ...JUUST ADD THAT DUPLICATE ROW IN ANOTHER TAB.................NAMES DUPLICATE DATA"
            row = [""] * len(HEADERS)
            if lookup_record is not None:
                for index, header in enumerate(HEADERS):
                    if header not in ("SUPPLIER_SLID", "NAME", "REF_NO", "REF_DT", "AMOUNT", "AMOUNT(SHEET1)", "Remark", "GST NO.", "PAN NUMBER", "TDS_NAME", "SERVICE_NAME"):
                        row[index] = lookup_record.get(header.lower(), "")
                row[IDX_SUPPLIER_SLID] = str(lookup_record.get("supplier_slid") or "").strip()
                row[IDX_NAME] = extract_vendor_name_from_invoice(record)
                row[IDX_AMOUNT_SHEET1] = lookup_record.get("amount", "")
                final_tds = str(lookup_record.get("tds_name") or norm_ai_tds).strip()
                final_service = str(lookup_record.get("service_name") or norm_ai_service).strip()
                row[IDX_TDS_NAME] = final_tds
                row[IDX_SERVICE_NAME] = final_service
                row[IDX_GST_NO] = clean_gst if clean_gst else lookup_record.get("gst number", "")
                row[IDX_PAN_NO] = clean_pan or raw_pan or str(lookup_record.get("pan number") or "")
            else:
                vendor_name = extract_vendor_name_fallback(record)
                row[IDX_SUPPLIER_SLID] = ""
                row[IDX_NAME] = vendor_name
                row[IDX_AMOUNT_SHEET1] = ""
                row[IDX_TDS_NAME] = norm_ai_tds or ai_tds
                row[IDX_SERVICE_NAME] = norm_ai_service or ai_service
                row[IDX_GST_NO] = clean_gst
                row[IDX_PAN_NO] = clean_pan or raw_pan or ""

            row[IDX_REF_NO] = inv_no
            row[IDX_REF_DT] = inv_date
            ai_amt = round_amount(record.get("AMOUNT(AI)"))
            row[IDX_AMOUNT] = ai_amt if (ai_amt not in (None, "")) else (round_amount(lookup_record.get("amount", "")) if lookup_record else "")

            ai_remark = str(record.get("Remark") or record.get("Remarks") or record.get("remark") or "").strip()
            dup_msg = f"DUPLICATE ROW (Matches Invoices row {existing_row_idx})"
            row[IDX_REMARK] = f"{ai_remark} | {dup_msg}".strip(" |")

            # Add to DUPLICATE DATA sheet with NORMAL styling (NOT RED)
            dup_target_row = get_last_data_row(dup_sheet) + 1
            for col_idx, val in enumerate(row, start=1):
                cell = dup_sheet.cell(row=dup_target_row, column=col_idx, value=val)
                cell.fill = PatternFill(fill_type=None)  # DON'T MARK RED
                cell.font = Font(color="000000")

            dup_sheet.cell(row=dup_target_row, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
            dup_sheet.cell(row=dup_target_row, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
            dup_sheet.cell(row=dup_target_row, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'

            # Unflag existing row in Invoices if it was marked red for duplicate
            if lookup_record is not None:
                for c in range(1, len(HEADERS) + 1):
                    sheet.cell(row=existing_row_idx, column=c).fill = PatternFill(fill_type=None)
                    sheet.cell(row=existing_row_idx, column=c).font = Font(color="000000")

            log.append([new_sync_key, datetime.now().isoformat()])
            print(
                f"Added DUPLICATE invoice {filename} (Ref: {inv_no}, GST: {clean_gst}) "
                f"to 'DUPLICATE DATA' tab at row {dup_target_row} (NOT marked red)."
            )
            synced_records.append(record)
            continue

        if lookup_record is None:
            # User request: "if gst no and pan number not match with sheet1 tab...then add that data also and marked that in red row"
            vendor_name = extract_vendor_name_fallback(record)
            row = [""] * len(HEADERS)
            row[IDX_SUPPLIER_SLID] = ""
            row[IDX_NAME] = vendor_name
            row[IDX_REF_NO] = inv_no
            row[IDX_REF_DT] = inv_date
            ai_amt = round_amount(record.get("AMOUNT(AI)"))
            row[IDX_AMOUNT] = ai_amt if (ai_amt not in (None, "")) else ""
            row[IDX_AMOUNT_SHEET1] = ""
            final_tds_name = norm_ai_tds or ai_tds
            row[IDX_TDS_NAME] = final_tds_name
            row[IDX_SERVICE_NAME] = norm_ai_service or ai_service
            row[IDX_GST_NO] = clean_gst
            row[IDX_PAN_NO] = clean_pan or raw_pan or ""

            is_multipage = bool(record.get("is_multipage") or tot_pages > 1)
            is_gst_missing = not bool(clean_gst)

            # No Sheet1 master row exists for this GST/PAN, so REF_ADMSITE_SHRTNAME,
            # TAG_ADMSITE_SHRTNAME, TAG_ADMSITE_AMOUNT, TDS, GST, and Paybale can't come
            # from the master directory. Estimate them instead from the invoice itself:
            # state name from the GST state code, site name from the AI-written remark,
            # and TDS/GST/Paybale from the statutory rate table.
            ai_remark_for_site = str(record.get("Remark") or record.get("Remarks") or record.get("remark") or "").strip()
            row[IDX_REF_ADMSITE_SHRTNAME] = derive_admsite_shrtname(clean_gst or raw_gst, clean_pan or raw_pan)
            row[IDX_TAG_ADMSITE_SHRTNAME] = extract_site_from_remark(ai_remark_for_site)
            row[IDX_TAG_ADMSITE_AMOUNT] = ai_amt if (ai_amt not in (None, "")) else ""
            est_tds, est_gst, est_paybale = estimate_tds_gst_paybale(ai_amt, final_tds_name)
            row[IDX_TDS] = est_tds
            row[IDX_GST] = est_gst
            row[IDX_PAYBALE] = est_paybale

            remarks_to_add = ["GST No. not available in Sheet1 (master directory) - TAG_ADMSITE/TDS/GST estimated from invoice"]
            if is_multipage:
                remarks_to_add.append(f"Multi-page PDF (Page {page_num} of {tot_pages})")
            if is_gst_missing:
                remarks_to_add.append("GST missing on invoice")

            ai_remark = str(record.get("Remark") or record.get("Remarks") or record.get("remark") or "").strip()
            all_remarks = []
            if ai_remark:
                all_remarks.append(ai_remark)
            if remarks_to_add:
                all_remarks.extend(remarks_to_add)
            row[IDX_REMARK] = " | ".join(all_remarks)

            target_row = get_last_data_row(sheet) + 1
            for col_idx, val in enumerate(row, start=1):
                cell = sheet.cell(row=target_row, column=col_idx, value=val)
                cell.fill = FLAGGED_FILL
                cell.font = FLAGGED_FONT

            sheet.cell(row=target_row, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=target_row, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=target_row, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=target_row, column=IDX_TDS + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=target_row, column=IDX_GST + 1).number_format = '"₹"#,##0.00'
            sheet.cell(row=target_row, column=IDX_PAYBALE + 1).number_format = '"₹"#,##0.00'

            log.append([new_sync_key, datetime.now().isoformat()])
            print(
                f"Added UNMATCHED invoice {filename} (page {page_num}/{tot_pages}) "
                f"({vendor_name or 'Unknown Vendor'}) [Flagged in RED: No Sheet1 Match] at row {target_row}."
            )
            synced_records.append(record)
            continue

        row = [""] * len(HEADERS)
        for index, header in enumerate(HEADERS):
            # Do NOT copy SUPPLIER_SLID, NAME, REF_NO, REF_DT, AMOUNT, AMOUNT(SHEET1), Remark, GST NO., PAN NUMBER, TDS_NAME, or SERVICE_NAME blindly from Sheet1
            if header not in ("SUPPLIER_SLID", "NAME", "REF_NO", "REF_DT", "AMOUNT", "AMOUNT(SHEET1)", "Remark", "GST NO.", "PAN NUMBER", "TDS_NAME", "SERVICE_NAME"):
                row[index] = lookup_record.get(header.lower(), "")

        # Set SUPPLIER_SLID from Sheet1 lookup aligned with GST NO.
        row[IDX_SUPPLIER_SLID] = str(lookup_record.get("supplier_slid") or "").strip()
        # Set vendor name from invoice ONLY (never from Sheet1)
        row[IDX_NAME] = extract_vendor_name_from_invoice(record)

        # Set AI extracted values
        row[IDX_REF_NO] = inv_no
        row[IDX_REF_DT] = inv_date
        ai_amt = round_amount(record.get("AMOUNT(AI)"))
        row[IDX_AMOUNT] = ai_amt if (ai_amt not in (None, "")) else round_amount(lookup_record.get("amount", ""))

        # Set Sheet1 amount based on matched record (GST/PAN + SERVICE_NAME + TDS_NAME)
        row[IDX_AMOUNT_SHEET1] = lookup_record.get("amount", "")

        # Determine final TDS_NAME and SERVICE_NAME from Sheet1 lookup
        final_tds = str(lookup_record.get("tds_name") or norm_ai_tds).strip()
        final_service = str(lookup_record.get("service_name") or norm_ai_service).strip()
        row[IDX_TDS_NAME] = final_tds
        row[IDX_SERVICE_NAME] = final_service

        # Set GST Number: prioritize invoice GST; fallback to master Sheet1 lookup
        row[IDX_GST_NO] = clean_gst if clean_gst else lookup_record.get("gst number", "")

        # Set PAN Number in its own column - never conflated with GST NO.
        row[IDX_PAN_NO] = clean_pan or raw_pan or str(lookup_record.get("pan number") or "")

        # Check flagging conditions
        is_multipage = bool(record.get("is_multipage") or tot_pages > 1)
        is_gst_missing = not bool(clean_gst)
        missing_ai_fields = get_missing_ai_fields(row)
        is_ai_missing = len(missing_ai_fields) > 0

        # Build descriptive remarks
        remarks_to_add = []
        if is_multipage:
            remarks_to_add.append(f"Multi-page PDF (Page {page_num} of {tot_pages})")
        if is_gst_missing:
            remarks_to_add.append("GST missing on invoice - matched via PAN")
        if is_ai_missing:
            remarks_to_add.append(f"Missing in invoice: {', '.join(missing_ai_fields)}")

        ai_remark = str(record.get("Remark") or record.get("Remarks") or record.get("remark") or "").strip()
        existing_remark = str(lookup_record.get("remarks", "")).strip()

        all_remarks = []
        if ai_remark:
            all_remarks.append(ai_remark)
        if existing_remark and existing_remark.upper() not in ("NONE", "GST") and existing_remark not in all_remarks:
            all_remarks.append(existing_remark)
        if remarks_to_add:
            all_remarks.extend(remarks_to_add)

        row[IDX_REMARK] = " | ".join(all_remarks)

        # Append directly at the very next available row to guarantee zero gaps
        target_row = get_last_data_row(sheet) + 1
        for col_idx, val in enumerate(row, start=1):
            cell = sheet.cell(row=target_row, column=col_idx, value=val)
            # Make row red ONLY if GST is missing or any (AI) column is missing; do NOT mark multi-page red
            if is_gst_missing or is_ai_missing:
                cell.fill = FLAGGED_FILL
                cell.font = FLAGGED_FONT

        sheet.cell(row=target_row, column=IDX_AMOUNT + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=target_row, column=IDX_AMOUNT_SHEET1 + 1).number_format = '"₹"#,##0.00'
        sheet.cell(row=target_row, column=IDX_TAG_ADMSITE_AMOUNT + 1).number_format = '"₹"#,##0.00'

        log.append([new_sync_key, datetime.now().isoformat()])
        flag_info = []
        if is_gst_missing:
            flag_info.append("GST missing")
        if is_ai_missing:
            flag_info.append(f"Missing AI: {', '.join(missing_ai_fields)}")
        flag_str = f" [Flagged in RED: {', '.join(flag_info)}]" if flag_info else ""

        print(
            f"Matched {filename} (page {page_num}/{tot_pages}) via {matched_via} "
            f"({lookup_record.get('name', 'Unknown Vendor')}){flag_str} at row {target_row}."
        )
        synced_records.append(record)

    if candidates or updated_existing > 0:
        atomic_save_workbook(workbook, EXCEL_PATH)

    if synced_records:
        for record in synced_records:
            record["excel_synced_at"] = datetime.now().isoformat()
        JSON_PATH.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    return len(synced_records) + updated_existing


def main() -> None:
    print(f"Watching {JSON_PATH} for new invoice records. Press Ctrl+C to stop.")
    while True:
        try:
            count = sync_once()
            if count:
                print(f"Synced {count} invoice record(s) to {EXCEL_PATH.name}.")
        except PermissionError:
            print("Excel file is open or locked. Close it; the watcher will retry.")
        except Exception as error:
            print(f"Sync error: {error}")
        time.sleep(POLL_SECONDS)


if __name__ == "__main__":
    main()
