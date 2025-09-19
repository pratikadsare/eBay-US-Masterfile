import io
import re
import sys
import copy
import datetime as dt
from typing import Dict, Tuple, List, Optional

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
try:
    # available in openpyxl >= 2.6
    from openpyxl.formula.translate import Translator
    HAS_TRANSLATOR = True
except Exception:
    HAS_TRANSLATOR = False


# ----------------------------
# CONFIG / CONSTANTS
# ----------------------------
st.set_page_config(page_title="ebay masterfile filler", layout="wide")

# Acceptable file types for uploads
ACCEPTED_TEMPLATE_TYPES = ("xlsx", "xls")
ACCEPTED_RAW_TYPES = ("xlsx", "xls", "csv")
ACCEPTED_MAP_TYPES = ("xlsx", "xls", "csv")

# ----------------------------
# UI HEADER (centered)
# ----------------------------
st.markdown("<h1 style='text-align:center; margin-bottom:0.2rem;'>ebay masterfile filler</h1>", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; font-size:1.05rem; opacity:0.85;'>Innovation in Action ⏐ Growth in Motion</div>", unsafe_allow_html=True)
st.divider()

# ----------------------------
# UTILS
# ----------------------------
def normalize(s) -> str:
    return str(s).strip().lower() if s is not None else ""

def read_any_dataframe(uploaded_file) -> pd.DataFrame:
    """Read CSV/XLS/XLSX into DataFrame with object dtype to preserve text (e.g., leading zeros)."""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        try:
            return pd.read_csv(uploaded_file, dtype=object, keep_default_na=False, na_values=[""])
        except UnicodeDecodeError:
            uploaded_file.seek(0)
            return pd.read_csv(uploaded_file, dtype=object, encoding="latin-1", keep_default_na=False, na_values=[""])
    else:
        # Excel: first sheet
        return pd.read_excel(uploaded_file, dtype=object)

def infer_mapping(df_map: pd.DataFrame) -> Dict[str, str]:
    """
    Build mapping {template_header -> raw_header} from a mapping file where
    the **FIRST column is RAW** and the **SECOND column is TEMPLATE**.
    Flexible header detection is supported; if ambiguous, we assume the first two columns
    are (raw, template) in that order.
    """
    if df_map is None or df_map.empty:
        raise ValueError("Mapping file appears empty. Provide at least two columns: RAW then TEMPLATE.")

    original_cols = list(df_map.columns)
    df_map.columns = [normalize(c) for c in df_map.columns]
    cols = list(df_map.columns)

    raw_keys = ("raw", "raw_header", "source", "raw_column", "input", "source_column", "rawsheet_column")
    template_keys = ("template", "template_header", "masterfile", "master", "target", "masterfile_column", "ebay", "ebay_column", "ebay_header")

    raw_col = next((c for c in cols if any(k in c for k in raw_keys)), None)
    template_col = next((c for c in cols if any(k in c for k in template_keys)), None)

    if raw_col is None or template_col is None:
        # Fallback strict order: first=raw, second=template
        if len(cols) < 2:
            raise ValueError(
                "Mapping file must have at least two columns (RAW first, TEMPLATE second). "
                f"Found columns: {original_cols}"
            )
        raw_col, template_col = cols[0], cols[1]

    mapping = {}
    for _, row in df_map.iterrows():
        raw_h = str(row[raw_col]).strip() if pd.notna(row[raw_col]) else ""
        tmpl_h = str(row[template_col]).strip() if pd.notna(row[template_col]) else ""
        if raw_h and tmpl_h and normalize(raw_h) != "nan" and normalize(tmpl_h) != "nan":
            # We want {template -> raw}
            mapping[tmpl_h] = raw_h
    if not mapping:
        raise ValueError("No valid mappings detected in mapping file (RAW first, TEMPLATE second).")
    return mapping

def get_template_headers(ws: Worksheet) -> Tuple[Dict[int, str], Dict[str, int], int]:
    """
    Read row-1 headers from the first sheet.
    Returns:
      - by_col: {col_idx -> header_text}
      - by_name: {normalized_header_text -> col_idx}
      - max_col: ws.max_column (for convenience)
    """
    max_col = ws.max_column
    by_col = {}
    by_name = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=1, column=c).value
        if v is not None and str(v).strip() != "":
            by_col[c] = str(v).strip()
            by_name[normalize(v)] = c
    return by_col, by_name, max_col

def is_cell_formula(cell) -> bool:
    v = cell.value
    if v is None:
        return False
    if hasattr(cell, "data_type") and getattr(cell, "data_type", None) == "f":
        return True
    if isinstance(v, str) and v.startswith("="):
        return True
    return False

def copy_style(src_cell, dst_cell):
    dst_cell.font = copy.copy(src_cell.font)
    dst_cell.fill = copy.copy(src_cell.fill)
    dst_cell.border = copy.copy(src_cell.border)
    dst_cell.alignment = copy.copy(src_cell.alignment)
    dst_cell.number_format = src_cell.number_format
    dst_cell.protection = copy.copy(src_cell.protection)

def copy_default_cell_to_row(ws: Worksheet, src_row: int, dst_row: int, col_idx: int):
    """Copy value (and style) from src_row to dst_row for a given column. Adjust formulas to dst row if possible."""
    src = ws.cell(row=src_row, column=col_idx)
    dst = ws.cell(row=dst_row, column=col_idx)

    # Style first (so number format etc. exists regardless of the value)
    copy_style(src, dst)

    if is_cell_formula(src) and isinstance(src.value, str):
        # Adjust formula to the destination row, if translator exists
        if HAS_TRANSLATOR:
            try:
                dst.value = Translator(src.value, origin=src.coordinate).translate_formula(dst.coordinate)
            except Exception:
                dst.value = src.value  # fallback: same formula (might keep row refs)
        else:
            dst.value = src.value
    else:
        dst.value = src.value

def nonempty(cell) -> bool:
    v = cell.value
    if v is None:
        return False
    if isinstance(v, str) and v.strip() == "":
        return False
    return True

def process_workbook(
    wb, raw_df: pd.DataFrame, mapping: Dict[str, str]
) -> Tuple[int, List[str]]:
    """
    Apply mapping into the FIRST sheet only.
    - Data filling begins from ROW 3.
    - Row 2 stays untouched.
    - Defaults in row 3 are preserved and copied down to each SKU row.
    - All other sheets remain untouched.
    Returns: (num_rows_written, skipped_columns_due_to_defaults)
    """
    ws: Worksheet = wb.worksheets[0]  # first sheet only

    # Headers from row-1
    by_col, by_name, max_col = get_template_headers(ws)

    # Identify protected columns (i.e., default exists in row 3 and must not be overwritten)
    protected_cols = set()
    defaults_present_cols = set()
    for c in range(1, max_col + 1):
        if nonempty(ws.cell(row=3, column=c)) or is_cell_formula(ws.cell(row=3, column=c)):
            protected_cols.add(c)
            defaults_present_cols.add(c)

    # Validate mapped template headers exist in row 1
    missing_template_headers = [h for h in mapping.keys() if normalize(h) not in by_name]
    if missing_template_headers:
        raise ValueError(
            "These template headers from your mapping do not exist in row‑1 of the masterfile: "
            + ", ".join(missing_template_headers)
        )

    # Validate mapped raw headers exist in raw_df columns (case-insensitive)
    raw_lookup = {normalize(c): c for c in raw_df.columns}
    missing_raw_headers = [r for r in mapping.values() if normalize(r) not in raw_lookup]
    if missing_raw_headers:
        raise ValueError(
            "These raw headers from your mapping are not present in your Raw sheet: "
            + ", ".join(missing_raw_headers)
        )

    # Build a filtered raw dataframe to count real rows (any mapped column having data)
    mapped_raw_cols = [raw_lookup[normalize(r)] for r in mapping.values()]
    candidate = raw_df[mapped_raw_cols].copy()
    # drop rows that are entirely empty across mapped columns
    non_empty_mask = ~(candidate.applymap(lambda x: (pd.isna(x) or str(x).strip() == "")).all(axis=1))
    data_df = raw_df[non_empty_mask].reset_index(drop=True)

    n_rows = len(data_df)
    if n_rows == 0:
        return 0, []

    # Precompute a list of (template_col_idx, raw_col_name)
    template_to_raw_pairs: List[Tuple[int, str]] = []
    for tmpl_hdr, raw_hdr in mapping.items():
        col_idx = by_name[normalize(tmpl_hdr)]
        raw_col = raw_lookup[normalize(raw_hdr)]
        template_to_raw_pairs.append((col_idx, raw_col))

    # Track which mapped template columns we skip because row-3 has defaults
    skipped_due_to_defaults = set()

    # Write rows starting at row=3
    start_row = 3
    for i in range(n_rows):
        target_row = start_row + i

        # 1) First, copy defaults from row 3 to target_row for ALL columns that have defaults
        if target_row > 3:
            for c in defaults_present_cols:
                copy_default_cell_to_row(ws, src_row=3, dst_row=target_row, col_idx=c)

        # 2) Then write mapped values for this row into non-protected columns only
        for col_idx, raw_col in template_to_raw_pairs:
            if col_idx in protected_cols:
                # do not overwrite defaults
                continue
            cell = ws.cell(row=target_row, column=col_idx)

            # Copy style from row 3 to keep formatting consistent even in non-default columns
            copy_style(ws.cell(row=3, column=col_idx), cell)

            val = data_df.iloc[i][raw_col]
            if pd.isna(val):
                val = None
            cell.value = val

    # Prepare skipped columns (names) for UI display
    by_col = {c: ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)}
    skipped_names = []
    for tmpl_hdr in mapping.keys():
        cidx = None
        key = normalize(tmpl_hdr)
        for c in range(1, ws.max_column + 1):
            v = ws.cell(row=1, column=c).value
            if v is not None and normalize(v) == key:
                cidx = c
                break
        if cidx is not None and cidx in protected_cols:
            skipped_names.append(by_col.get(cidx, f"COL{cidx}"))

    return n_rows, sorted(set(skipped_names))


# ----------------------------
# INPUT SECTIONS (3 required)
# ----------------------------
st.subheader("1) Upload Template (XLS/XLSX)")
template_file = st.file_uploader("Upload the eBay masterfile template", type=ACCEPTED_TEMPLATE_TYPES, key="template_upload")

st.subheader("2) Upload Raw sheet (CSV/XLSX)")
raw_file = st.file_uploader("Upload Raw sheet", type=ACCEPTED_RAW_TYPES, key="raw_upload")

st.subheader("3) Upload Mapping file (first column = RAW, second column = TEMPLATE)")
mapping_file = st.file_uploader("Upload Mapping file", type=ACCEPTED_MAP_TYPES, key="map_upload")

with st.expander("Mapping file guidance (updated order)", expanded=False):
    st.markdown(
        """
        **Required order:** the **first column is RAW** and the **second column is TEMPLATE**.  
        Flexible names are supported, but order matters when headers are ambiguous.

        **Example CSV:**
        ```
        raw,template
        Sku,SKU
        Product Title,Title
        Price,Price
        Qty,Quantity
        ```
        - Mapping is matched to **row‑1 headers** of the template.  
        - Data writing begins at **row 3** (row 2 remains untouched).  
        - Any **defaults in row 3** (e.g., Site ID, Currency) are **preserved** and copied down for each SKU row.  
        """
    )

# ----------------------------
# PROCESS
# ----------------------------
if template_file and raw_file and mapping_file:
    try:
        # 1) Read user uploads
        raw_df = read_any_dataframe(raw_file)
        map_df = read_any_dataframe(mapping_file)
        mapping = infer_mapping(map_df)

        # 2) Load workbook from uploaded template (first sheet to be updated; others preserved)
        template_bytes = template_file.read()
        wb = load_workbook(filename=io.BytesIO(template_bytes), data_only=False)  # keep formulas intact

        # 3) Process workbook
        num_rows, skipped_cols = process_workbook(wb, raw_df, mapping)

        if num_rows == 0:
            st.warning("No data rows found in the Raw sheet (all mapped columns were empty).")
        else:
            st.success(f"Filled {num_rows} row(s) into the first sheet starting at row 3.")
            if skipped_cols:
                st.info(
                    "These mapped template columns were **not** overwritten because row‑3 contains defaults: "
                    + ", ".join(skipped_cols)
                )

        # 4) Save to buffer and offer download
        out_buf = io.BytesIO()
        wb.save(out_buf)
        out_buf.seek(0)

        today_str = dt.date.today().isoformat()  # YYYY-MM-DD
        download_name = f"ebay_masterfile_filled_{today_str}.xlsx"
        st.download_button(
            label="⬇️ Download filled masterfile",
            data=out_buf.getvalue(),
            file_name=download_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        # Optional preview (first 5 of raw)
        with st.expander("Preview: first 5 rows of your Raw sheet"):
            st.dataframe(raw_df.head(5))

    except Exception as e:
        st.error(f"Processing failed: {e}")
        st.exception(e)
else:
    st.info("Please upload the **Template**, **Raw sheet**, and **Mapping file** to begin.")
