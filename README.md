# GSTR-1 HSN Calculator

Desktop app to consolidate HSN-wise data from bills (JPG, PDF, Excel) for GSTR-1 filing.

---

## Setup (one time only)

### Step 1 — Install dependencies
Open Command Prompt and run:

```
pip install customtkinter anthropic openpyxl PyMuPDF Pillow
```

### Step 2 — Run the app

```
python app.py
```

---

## Usage

1. **Enter your Anthropic API key** in the sidebar and click Save Key
   - Get your key at: https://console.anthropic.com
   - Key is saved locally on your PC — never leaves your machine

2. **Upload bills** — click Browse Files and select JPG / PNG / PDF / XLSX / CSV files

3. **Click Process Bills** — the app reads each file, extracts HSN data via AI, and consolidates

4. **Review the table** — double-click any cell to edit if something was extracted incorrectly

5. **If HSN is missing** — a warning banner appears, click Enter HSN to assign the code

6. **Download Excel** — exports GSTR1_HSN_Summary.xlsx ready to upload to the portal

---

## What gets extracted per line item

| Field | Description |
|-------|-------------|
| HSN Code | As printed on the bill (or UNKNOWN if missing) |
| UQC | Unit (NOS, KGS, MTR, LTR, etc.) |
| Quantity | Numeric quantity |
| Taxable Value | Assessable/taxable amount |
| IGST | Interstate tax (0 if intrastate) |
| CGST | Central tax (0 if interstate) |
| SGST | State tax (0 if interstate) |

Same HSN codes across multiple bills are **automatically summed**.

---

## Excel column header detection

The app auto-detects these Excel column name variations:

- **HSN**: `HSN`, `HSN Code`, `HSN/SAC`, `SAC`, `SAC Code`
- **Quantity**: `Quantity`, `Qty`, `Units`
- **Taxable Value**: `Taxable Value`, `Taxable Amount`, `Assessable Value`
- **IGST**: `IGST`, `IGST Amount`, `Integrated Tax`
- **CGST**: `CGST`, `CGST Amount`, `Central Tax`
- **SGST**: `SGST`, `SGST Amount`, `State Tax`, `UTGST`

---

## Notes

- API key is stored in `~/.gstr1_config.json` on your PC
- Large PDFs are processed page by page
- Handwritten invoices may have lower accuracy
- Always review the table before downloading
