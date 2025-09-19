# eBay Masterfile Processor (Streamlit)

## Quick Start
1. Create a virtual env and install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the app:
   ```bash
   streamlit run streamlit_app.py
   ```

## What this app does
- Always fetches the latest version of your fixed Google Sheet template (first sheet only).
- Users upload exactly two files: a **Raw sheet** and a **Mapping file**.
- Mapping is matched to **row‑1 headers**, data writing starts at **row 3**, row 2 stays untouched.
- Preserves defaults in row‑3 (Site ID, Currency, etc.) and **copies them down** for each SKU row.
- **Does not touch other sheets** (names, formatting, formulas, values remain intact).
- Lets the user download the filled masterfile as an Excel file named `ebay_masterfile_filled_YYYY-MM-DD.xlsx`.

## Mapping file format (flexible)
Two columns minimum:
- One for the template column (e.g., `template`, `masterfile`, `ebay_column`)
- One for the raw column (e.g., `raw`, `raw_header`, `source`)

Example:
```csv
template,raw
SKU,Sku
Title,Product Title
Price,Price
Quantity,Qty
```

> Ensure the Google Sheet is shared as **Anyone with the link: Viewer**.
