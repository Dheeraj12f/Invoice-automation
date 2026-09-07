"""Continuously copy newly saved invoice JSON records into the Excel ledger."""

from __future__ import annotations

import json
import os
import re
import shutil
import time
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE_DIR = Path(__file__).resolve().parent
JSON_PATH = BASE_DIR / "extracted_invoices.json"
EXCEL_PATH = BASE_DIR / "invoices" / "invoice_ledger.xlsx"
POLL_SECONDS = 2
SHEET1_HEADER_ROW = 2

HEADERS = [
    "S.No", "invoice no.(Ai)", "invoice date(AI)", "SRVDT", "SUPPLIER_SLID",
    "NAME", "REF_ADMSITE_SHRTNAME", "TDS_NAME", "REF_NO", "REF_DT",
    "SERVICE_NAME", "AMOUNT(AI)", "AMOUNT", "TAG_ADMSITE_SHRTNAME",
    "Source Site", "TAG_ADMSITE_AMOUNT", "TDS", "GST", "Paybale", "TERM",
    "FORM_NAME", "DESCRIPATION", "Remarks", "PAN NUMBER", "GST NUMBER",
    "GST IN Status", "GST IN HOLD AMOUNT", "LL email id", "LL email id-2",
]

FLAGGED_FILL = PatternFill("solid", fgColor="FFD9D9")
FLAGGED_FONT = Font(color="9C0006")


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


def load_sheet1_lookup(
    workbook,
) -> tuple[dict[str, dict[str, object]], dict[str, dict[str, object]]]:
    """Return Sheet1 rows indexed by both GST NUMBER and PAN NUMBER values."""
    if "Sheet1" not in workbook.sheetnames:
        raise RuntimeError(
            "Sheet1 is missing. Add the data-validation tab named Sheet1 "
            "before syncing invoice records."
        )

    sheet = workbook["Sheet1"]
    source_headers = [
        str(cell.value).strip() if cell.value is not None else ""
        for cell in sheet[SHEET1_HEADER_ROW]
    ]
    header_positions = {
        header.lower(): index for index, header in enumerate(source_headers) if header
    }
    gst_position = header_positions.get("gst number")
    pan_position = header_positions.get("pan number")
    if gst_position is None and pan_position is None:
        raise RuntimeError(
            f"Sheet1 must have a 'GST NUMBER' or 'PAN NUMBER' header in row {SHEET1_HEADER_ROW}."
        )

    gst_lookup: dict[str, dict[str, object]] = {}
    pan_lookup: dict[str, dict[str, object]] = {}
    for values in sheet.iter_rows(min_row=SHEET1_HEADER_ROW + 1, values_only=True):
        row_dict = {
            header.lower(): values[index] if index < len(values) else ""
            for index, header in enumerate(source_headers) if header
        }
        if gst_position is not None and gst_position < len(values):
            gst_number = normalise_gst_number(values[gst_position])
            if gst_number:
                gst_lookup[gst_number] = row_dict
        if pan_position is not None and pan_position < len(values):
            pan_number = normalise_pan_number(values[pan_position])
            if pan_number:
                pan_lookup[pan_number] = row_dict

    return gst_lookup, pan_lookup


def get_last_data_row(sheet) -> int:
    """Return the 1-based index of the last row containing real non-empty data."""
    for row_idx in range(sheet.max_row, 1, -1):
        if any(sheet.cell(row=row_idx, column=c).value not in (None, "") for c in range(1, len(HEADERS) + 1)):
            return row_idx
    return 1


def get_missing_ai_fields(row: list) -> list[str]:
    """Check if any header ending in (AI)/(Ai) has a missing, unreadable, or invalid value."""
    missing = []
    # Check invoice no.(Ai) (column index 1)
    inv_no = str(row[1]).strip() if row[1] is not None else ""
    if not inv_no or inv_no.upper() in {"N/A", "NA", "NONE", "NOT FOUND", "NOT LEGIBLE", "OCR UNCLEAR", "NULL", "-"}:
        missing.append("invoice no.(Ai)")

    # Check invoice date(AI) (column index 2)
    inv_date = str(row[2]).strip() if row[2] is not None else ""
    if not inv_date or inv_date.upper() in {"N/A", "NA", "NONE", "NOT FOUND", "NOT LEGIBLE", "NULL", "-"}:
        missing.append("invoice date(AI)")

    # Check AMOUNT(AI) (column index 11)
    amt = row[11]
    if amt is None or str(amt).strip() in {"", "None", "null", "N/A", "NA", "NONE"}:
        missing.append("AMOUNT(AI)")
    else:
        try:
            val_num = float(str(amt).replace(",", "").replace("₹", "").strip())
            if val_num <= 0:
                missing.append("AMOUNT(AI)")
        except ValueError:
            missing.append("AMOUNT(AI)")

    return missing


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

            # Also check if any AI fields are missing in this existing row
            missing_ai = get_missing_ai_fields(vals)
            if missing_ai:
                is_red = True
                curr_remark = str(vals[22]).strip() if vals[22] else ""
                miss_str = f"Missing in invoice: {', '.join(missing_ai)}"
                if miss_str not in curr_remark:
                    if curr_remark and curr_remark.upper() != "NONE":
                        vals[22] = f"{curr_remark} | {miss_str}"
                    else:
                        vals[22] = miss_str

            data_rows.append((vals, is_red))

    # Check if there is a gap (i.e. if max_row != 1 + len(data_rows))
    last_row = get_last_data_row(sheet)
    has_gaps = (last_row != 1 + len(data_rows)) or any(
        all(sheet.cell(row=r, column=c).value in (None, "") for c in range(1, len(HEADERS) + 1))
        for r in range(2, last_row + 1)
    )

    if not has_gaps:
        return

    # Clear everything from row 2 downwards
    sheet.delete_rows(2, sheet.max_row)

    # Re-write the data rows compactly starting at row 2
    for idx, (vals, is_red) in enumerate(data_rows, start=2):
        for c, val in enumerate(vals, start=1):
            cell = sheet.cell(row=idx, column=c, value=val)
            if is_red:
                cell.fill = FLAGGED_FILL
                cell.font = FLAGGED_FONT
        sheet.cell(row=idx, column=12).number_format = '"₹"#,##0.00'


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


def ensure_ledger():
    """Open the ledger, creating the expected invoice sheet when needed."""
    EXCEL_PATH.parent.mkdir(parents=True, exist_ok=True)
    workbook = load_ledger_safely(EXCEL_PATH)
    sheet = workbook["Invoices"] if "Invoices" in workbook.sheetnames else workbook.active
    headers = [cell.value for cell in sheet[1]] if sheet.max_row else []
    if headers != HEADERS:
        if headers and any(headers):
            sheet.title = "Invoices_legacy"
            sheet = workbook.create_sheet("Invoices", 0)
        else:
            sheet.title = "Invoices"
            sheet.delete_rows(1, sheet.max_row)
        sheet.append(HEADERS)
        for column in range(1, len(HEADERS) + 1):
            cell = sheet.cell(row=1, column=column)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="1F4E78")
            cell.alignment = Alignment(horizontal="center")
            sheet.column_dimensions[cell.column_letter].width = 20
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
    else:
        # Compact any empty gap rows between header row 1 and data rows
        compact_invoice_sheet(sheet)

    if "SyncLog" not in workbook.sheetnames:
        log = workbook.create_sheet("SyncLog")
        log.append(["sync_key", "synced_at"])
        log.sheet_state = "hidden"
    return workbook, sheet, workbook["SyncLog"]


def sync_once() -> int:
    """Write each unsynced saved JSON record exactly once to Excel."""
    if not JSON_PATH.exists():
        return 0
    try:
        payload = json.loads(JSON_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return 0  # The server is still writing the JSON; retry on the next poll.

    candidates = [
        record for record in payload.get("invoices", [])
        if record.get("status") == "saved" and not record.get("excel_synced_at")
    ]
    if not candidates:
        return 0

    workbook, sheet, log = ensure_ledger()
    gst_lookup, pan_lookup = load_sheet1_lookup(workbook)
    logged_keys = {row[0] for row in log.iter_rows(min_row=2, values_only=True) if row[0]}
    synced_records = []
    for record in candidates:
        page_num = record.get("page_number", 1)
        tot_pages = record.get("total_pages", 1)
        inv_no = record.get("invoice no.(Ai)", "")
        saved_at = record.get("saved_at", "")
        filename = record.get("filename", "")

        old_sync_key = f"{filename}|{saved_at}"
        new_sync_key = f"{filename}|page_{page_num}|inv_{inv_no}|{saved_at}"

        if old_sync_key not in logged_keys and new_sync_key not in logged_keys:
            raw_gst = record.get("GST NUMBER", "")
            raw_pan = record.get("PAN NUMBER", "")

            clean_gst = normalise_gst_number(raw_gst)
            clean_pan = normalise_pan_number(raw_pan) or normalise_pan_number(raw_gst)

            lookup_record = None
            matched_via = None
            if clean_gst and clean_gst in gst_lookup:
                lookup_record = gst_lookup[clean_gst]
                matched_via = f"GST {clean_gst}"
            elif clean_pan and clean_pan in pan_lookup:
                lookup_record = pan_lookup[clean_pan]
                matched_via = f"PAN {clean_pan}"

            if lookup_record is None:
                print(
                    f"Skipped {filename} (page {page_num}/{tot_pages}): no Sheet1 row "
                    f"matches GST NUMBER {raw_gst!r} or PAN NUMBER {raw_pan!r}."
                )
                continue

            row = [""] * len(HEADERS)
            for index, header in enumerate(HEADERS):
                row[index] = lookup_record.get(header.lower(), "")
            row[1] = record.get("invoice no.(Ai)", "")
            row[2] = record.get("invoice date(AI)", "")
            row[11] = record.get("AMOUNT(AI)", "")

            # If invoice had a specific GST, use it; otherwise retain GST from Sheet1 lookup
            if clean_gst:
                row[24] = clean_gst
            # If invoice had a specific PAN, use it; otherwise retain PAN from Sheet1 lookup
            if clean_pan:
                row[23] = clean_pan

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

            if remarks_to_add:
                existing_remark = str(row[22]).strip() if row[22] else ""
                if existing_remark and existing_remark.upper() != "NONE":
                    row[22] = f"{existing_remark} | {'; '.join(remarks_to_add)}"
                else:
                    row[22] = "; ".join(remarks_to_add)

            # Append directly at the very next available row to guarantee zero gaps
            target_row = get_last_data_row(sheet) + 1
            for col_idx, val in enumerate(row, start=1):
                cell = sheet.cell(row=target_row, column=col_idx, value=val)
                # Make row red if multi-page invoice, GST is missing, or any (AI) column is missing
                if is_multipage or is_gst_missing or is_ai_missing:
                    cell.fill = FLAGGED_FILL
                    cell.font = FLAGGED_FONT

            sheet.cell(row=target_row, column=12).number_format = '"₹"#,##0.00'

            log.append([new_sync_key, datetime.now().isoformat()])
            flag_info = []
            if is_multipage:
                flag_info.append("Multi-page")
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
    atomic_save_workbook(workbook, EXCEL_PATH)

    for record in synced_records:
        record["excel_synced_at"] = datetime.now().isoformat()
    JSON_PATH.write_text(json.dumps(payload, indent=2) + "\n", encoding="utf-8")
    return len(synced_records)


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
