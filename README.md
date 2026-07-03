# GSTR-1 HSN Summary Calculator

Desktop app that consolidates HSN-wise data from purchase/sales bills (images, PDFs, Excel) into a GSTR-1-ready HSN Summary in Excel.

Invoice images and PDFs are read using **Claude Haiku (Vision)** via the Anthropic API. Excel/CSV files are parsed locally at no cost — no API key needed if you only process spreadsheets.

---

## Features

- **Multi-format input** — JPG / PNG / PDF (AI Vision) and XLSX / XLS / CSV (parsed directly, free)
- **Automatic HSN consolidation** — identical HSN codes across all bills are summed into one row
- **Fully editable results** — double-click any cell to correct AI extraction before export
- **UNKNOWN-HSN handling** — items missing an HSN code are flagged in a banner; assign the code in one dialog and totals merge automatically
- **Formatted Excel export** — styled `GSTR1_HSN_Summary.xlsx` with SUM totals row and frozen header
- **Color-coded processing log** — green = success, red = error, amber = warning
- **Per-file queue management** — remove individual files without clearing the whole queue
- **Light / dark mode** toggle

---

## Setup

### 1. Clone and create a virtual environment

```powershell
git clone https://github.com/shetty30/hsn_summary_calculator.git
cd hsn_summary_calculator
python -m venv venv
.\venv\Scripts\Activate
```

### 2. Install dependencies

```powershell
pip install -r requirements.txt
```

### 3. Run

```powershell
python app.py
```

---

## Usage

1. **Enter your Anthropic API key** in the sidebar and click *Save Key*
   - Get a key at https://console.anthropic.com
   - Stored locally in `~/.gstr1_config.json` — never leaves your machine
   - **Not required** if you only process Excel/CSV files
2. **Upload bills** — *Browse Files*, select any mix of JPG / PNG / PDF / XLSX / CSV
3. **Process Bills** — each image/PDF page is sent to Claude Vision; spreadsheets are parsed locally
4. **Review** — check the stat cards and table; double-click cells to fix any extraction errors
5. **Assign missing HSNs** — if a warning banner appears, click *Enter HSN →*
6. **Download Excel** — export the consolidated summary, ready for the GST portal

---

## What gets extracted per line item

| Field | Description |
|-------|-------------|
| HSN Code | As printed on the bill (`UNKNOWN` if missing) |
| UQC | Unit (NOS, KGS, MTR, LTR, etc.) |
| Quantity | Numeric quantity |
| Taxable Value | Assessable/taxable amount (INR) |
| IGST | Interstate tax (0 if intrastate) |
| CGST | Central tax (0 if interstate) |
| SGST / UTGST | State tax (0 if interstate) |

---

## Excel column header detection

The parser auto-detects common column name variations:

- **HSN**: `HSN`, `HSN Code`, `HSN/SAC`, `SAC`, `SAC Code`
- **Quantity**: `Quantity`, `Qty`, `Units`, `Nos`
- **UQC**: `UQC`, `Unit`, `UOM`
- **Taxable Value**: `Taxable Value`, `Taxable Amount`, `Assessable Value`, `Basic Amount`
- **IGST**: `IGST`, `IGST Amount`, `Integrated Tax`
- **CGST**: `CGST`, `CGST Amount`, `Central Tax`
- **SGST**: `SGST`, `SGST Amount`, `State Tax`, `UTGST`

---

## Building a standalone .exe (optional)

For machines without Python installed:

```powershell
pip install pyinstaller
pyinstaller --onefile --windowed --collect-all customtkinter --name "HSN_Calculator" app.py
```

Output: `dist\HSN_Calculator.exe`. Notes:

- `--collect-all customtkinter` is required — the exe crashes without its bundled theme assets
- First launch takes ~5–10 s (self-extracting) — this is normal
- Windows SmartScreen may warn on unsigned exes: *More info → Run anyway*

---

## Accuracy notes

- AI extraction from images is good but **not guaranteed** — dense multi-column tax tables and low-resolution scans can produce swapped CGST/SGST values or misread quantities
- **Always verify** that the summary's taxable value + tax totals tie back to the invoice grand totals before filing
- Handwritten invoices have lower accuracy
- Large PDFs are processed page by page

---

## Tech stack

Python · CustomTkinter · Anthropic API (Claude Haiku Vision) · PyMuPDF · openpyxl
