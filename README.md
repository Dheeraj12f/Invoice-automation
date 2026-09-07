# Invoice Automation MCP Server & Excel Ledger Sync

An automated invoice processing pipeline built on the **Model Context Protocol (MCP)**. This system allows multimodal AI models (such as Claude Desktop) to discover, inspect, and extract key metadata from invoice files—including PDFs, Word documents (`.docx`), Excel spreadsheets (`.xlsx`), and smartphone scan images—and automatically synchronize and enrich the extracted data with a master vendor directory in an Excel ledger.

---

## 🌟 Key Features

1. **Multimodal MCP Server (`server.py`)**:
   - **Supported Formats**: Scanned & digital PDFs (`.pdf`), Word documents (`.docx`), Excel spreadsheets (`.xlsx`), and images (`.png`, `.jpg`, `.jpeg`).
   - **Large File Optimization (>1MB)**: Automatically downscales and compresses high-resolution photos and scanned PDF pages to high-quality JPEG under 900 KB, preventing MCP timeouts and buffer drops.
   - **Multi-Page Invoice Handling**: Reads and extracts individual pages from multi-page PDFs, safely retaining the inbox file until all pages are processed.
   - **Safe File Deletion**: Automatically deletes processed files from the inbox once extraction is complete.

2. **Excel Synchronization Engine (`sync_json_to_excel.py`)**:
   - **Master Vendor Lookup**: Matches invoices by GSTIN or PAN against `Sheet1` to enrich records with Supplier ID, vendor names, site names, TDS terms, and email addresses.
   - **Fallback to PAN**: If GSTIN is absent or unregistered on an invoice, automatically falls back to matching via PAN.
   - **Conditional Red Row Styling**: Highlights rows in red (`#FFD9D9` fill with `#9C0006` text) if:
     - The invoice comes from a multi-page PDF.
     - The GST number is missing.
     - Any AI-extracted column (`invoice no.(Ai)`, `invoice date(AI)`, `AMOUNT(AI)`) is missing or unreadable.
   - **Descriptive Remarks**: Records explanatory notes in the `Remarks` column (e.g. `Multi-page PDF (Page 1 of 2); Missing in invoice: invoice no.(Ai)`).
   - **Zero Gap Compaction & Atomic Saving**: Prevents and removes blank gap rows, and saves atomically with rolling `.bak` backups to prevent file corruption.

---

## 📁 Repository Structure

```
.
├── invoices/
│   ├── inbox/                # Input folder for incoming invoices (PDF, DOCX, XLSX, JPG, PNG)
│   ├── failed/               # Fallback folder for failed files
│   ├── processed/            # Compatibility directory
│   └── invoice_ledger.xlsx   # Main Excel workbook (Invoices, Sheet1 master, SyncLog)
├── extracted_invoices.json   # Intermediate JSON store populated by the MCP server
├── requirements.txt          # Python dependencies
├── server.py                 # MCP Server for invoice reading and extraction storage
└── sync_json_to_excel.py     # Background watcher syncing JSON records into Excel
```

---

## 🚀 Setup & Installation

### 1. Install Dependencies
```bash
pip install -r requirements.txt
```

### 2. Configure Claude Desktop (MCP)
Add the server to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "invoice-server": {
      "command": "python",
      "args": [
        "D:\\User profile\\89162\\Downloads\\Invoice automation\\invoice-mcp-server\\server.py"
      ]
    }
  }
}
```

### 3. Run the Excel Sync Watcher
Run the background synchronizer in a separate terminal:

```bash
python sync_json_to_excel.py
```

---

## 🛠️ MCP Tools Exposed

* `list_inbox_invoices()`: Lists all pending invoice documents in `invoices/inbox`.
* `read_invoice_content(filename, page_number=1)`: Reads text from Word/Excel/digital PDFs, or returns optimized visual image content for scanned pages/photos.
* `save_extracted_data_to_json(...)`: Stores extracted metadata (`invoice no.(Ai)`, `invoice date(AI)`, `AMOUNT(AI)`, `GST NUMBER`, `PAN NUMBER`, page information) and safely unlinks the source file when finished.
* `get_extracted_invoices()`: Retrieves all stored JSON records.
"# Invoice-automation" 
