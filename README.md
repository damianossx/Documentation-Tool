# SmartDocs Insight

> Desktop assistant for analyzing Rockwell invoice CSV/PDF files, extracting COO and PO/Box 5 data, and generating Excel tracking sheets for Certificate of Origin requests.

![SmartDocs Insight UI](docs/images/ui-overview.png)

---

## ✨ Key Features

- **Invoice & COO analysis**
  - Parses Rockwell invoice CSV files and linked PDF documents.
  - Extracts non‑EU and EU items based on COO codes.
  - Calculates total non‑EU weight in kilograms.

- **Missing COO detection**
  - Scans line items for incomplete or malformed COO/weight information.
  - Shows a dedicated, resizable modal listing all lines that require manual review.

- **PO / Box 5 and invoice reference extraction**
  - Collects PO / Box 5 numbers and invoice IDs from CSV header and items.
  - Builds a clean summary that can be used in classification / CoO request emails.

- **Excel tracking sheet export**
  - Generates an Excel file with one line per PO/SO pair using `openpyxl`.
  - Automatically fills key metadata (customer name, BPID from PDF, dates, status, responsible person).
  - Optionally opens the generated file in Excel.

- **Email body generator**
  - Prepares a structured email body with:
    - Non‑EU item list,
    - EU item list,
    - Reference invoice(s),
    - Total non‑EU weight and Box 5 references.

- **User‑friendly GUI**
  - Built with `tkinter`.
  - Dark/Light mode toggle with preference persisted to `user_preferences.json`.
  - Status panel with processing duration and total invoice count.
  - Top‑left modals to avoid covering other applications.

- **Logging & diagnostics**
  - Writes technical logs to `SmartDocs.log`.
  - Supports a debug mode via environment variables.

---

## 🧩 Architecture Overview

The core logic lives in `smartdocs_insight/main.py` and is organized into several functional blocks:

- **Invoice analysis**
  - `analyze_invoice(file_path)`  
    Parses a Rockwell invoice CSV file and returns:
    - Non‑EU items
    - EU items
    - Total non‑EU weight (kg)
    - Box 5 reference
    - Invoice reference
    - Missing COO alerts

- **Metadata extraction**
  - `extract_csv_metadata(file_path)`  
    Reads invoice number, customer name and PO/SO pairs from the CSV.
  - `extract_bpid_from_pdf(file_path)`  
    Opens the PDF using `PyMuPDF` (`fitz`) and extracts the BPID / customer ID based on localized labels.

- **Export & clipboard utilities**
  - `run_metadata_export(file_paths, responsible_person, open_excel)`  
    Builds an Excel tracking file using `openpyxl` and prepares tab‑separated data for clipboard.

- **GUI**
  - `show_gui()`  
    Initializes the main window:
    - file selection,
    - dark/light mode toggle,
    - email output text area,
    - status bar with invoice analysis counter.
  - `get_responsible_person()`  
    Modal dialog for selecting or typing the responsible person and choosing whether to auto‑open Excel.

- **Preferences & counters**
  - Dark mode, invoice analysis counter, and button color are stored in `user_preferences.json` in the same folder as the script.

For a more detailed breakdown of the main functions and flows, see [`docs/architecture.md`](docs/architecture.md).

---

## 📦 Installation

### 1. Clone the repository

```bash
git clone https://github.com/<your-github-username>/smartdocs-insight.git
cd smartdocs-insight