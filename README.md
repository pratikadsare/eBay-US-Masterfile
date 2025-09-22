# ebay masterfile filler (Streamlit)

**Tagline:** *Innovation in Action ⏐ Growth in Motion*

## Quick Start
1. Create a virtual env and install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the app:
   ```bash
   streamlit run streamlit_app.py
   ```

## What this app does (v3 updates)
- Users upload **three** inputs: the **Template (XLS/XLSX)**, a **Raw sheet (CSV/XLSX)**, and a **Mapping file**.
- **Mapping file order:** Column 1 = **TEMPLATE**, Column 2 = **RAW**. Mapping matches **row‑1 headers** in the Template.
- **Row 2 now serves as the defaults row** (the previous row 3 behavior moved up). Writing starts at **row 2**.
- Defaults present in row‑2 are preserved and **copied down** for every SKU row.
- Only the **first sheet** is updated; **all other sheets remain untouched** (names, formatting, formulas, values).

## Mapping file example
```csv
template,raw
SKU,Sku
Title,Product Title
Price,Price
Quantity,Qty
```
