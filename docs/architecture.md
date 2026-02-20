# SmartDocs Insight – Architecture

## 1. High-Level Flow

1. User selects one or more CSV + PDF files for a single invoice.
2. The application validates:
   - CSV/PDF filename signatures to ensure they belong to the same invoice.
   - Consistent BPID across all selected PDFs.
3. For each CSV file:
   - Parse header and line items.
   - Extract COO and weight data from the line item text.
   - Classify items into EU / non‑EU based on country code.
   - Collect Box 5 / PO references and invoice numbers.
   - Collect missing COO lines and file issues.
4. Aggregated results are used to:
   - Build a structured email body for CoO requests.
   - Optionally generate an Excel tracking sheet for internal tracking.

## 2. Main Components

- **Invoice analysis (`analyze_invoice`)**
  - Single‑file CSV parser.
  - Returns EU/non‑EU items, total non‑EU weight, Box 5 reference, invoice reference and missing COO alerts.

- **Metadata extraction**
  - `extract_csv_metadata`:
    - Reads customer name, invoice number and PO/SO pairs.
  - `extract_bpid_from_pdf`:
    - Uses `PyMuPDF` (`fitz`) to read PDF text and extract BPID/customer ID from localized labels.

- **Export & clipboard**
  - `run_metadata_export`:
    - Builds an Excel workbook using `openpyxl`.
    - Stores a row per unique PO/SO pair.
    - Copies a tab‑separated summary to the clipboard via `pyperclip`.

- **GUI layer (Tkinter)**
  - `show_gui`:
    - Main window setup (theme, size, layout).
    - File selection dialog and validation logic.
    - Email preview text area and status panel.
  - `get_responsible_person`:
    - Modal dialog for selecting or entering the responsible person.
    - Option to open the generated Excel immediately.

- **Preferences & counters**
  - `user_preferences.json`:
    - Stores dark mode state, usage counters and optional accent color.
  - Helper functions to increment and reset the invoice analysis counter.

## 3. Data Inputs & Outputs

**Inputs**

- Invoice CSV files exported from Rockwell systems.
- Invoice PDF files containing customer ID (BPID) in multiple language variants.

**Outputs**

- Email body (plain text) with:
  - Non‑EU item list,
  - EU item list,
  - Reference invoice(s),
  - Total non‑EU weight and Box 5 reference.
- Excel tracking sheet:
  - One row per PO/SO pair.
  - Includes customer data, BPID, dates, status and responsible person.
- Log entries:
  - Written to `SmartDocs.log`.

## 4. Error Handling & Alerts

- Validation errors (mismatched invoice signatures, inconsistent BPIDs, missing file types) are surfaced via message boxes.
- Missing COO lines and file issues are displayed in dedicated, resizable modals:
  - Each modal supports "Copy All" for quick sharing.

## 5. Future Enhancements (Proposed)

- Configurable country code lists (EU vs non‑EU).
- CLI mode for unattended/batch runs.
- Automated test suite (Pytest) around core parsing functions.
- Optional HTML/PDF summary export for audit trails.