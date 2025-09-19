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

## What this app does
- Users upload **three** inputs: the **Template (XLS/XLSX)**, a **Raw sheet (CSV/XLSX)**, and a **Mapping file**.
- Mapping file order is **RAW first, TEMPLATE second**; mapping is matched to **row‑1 headers** of the Template.
- Writing starts at **row 3**; row 2 stays untouched.
- Defaults present in row‑3 are preserved and **copied down** for every SKU row.
- Only the **first sheet** is updated; **all other sheets remain untouched** (names, formatting, formulas, values).

## Mapping file example
```csv
raw,template
Sku,SKU
Product Title,Title
Price,Price
Qty,Quantity
```
