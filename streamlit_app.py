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

HEADERS_ROW = 1         # Row 1 = headers (mapping is done using these)
DEFAULTS_SRC_ROW = 2    # Row 2 = default values row (copied down per SKU)
START_WRITE_ROW = 2     # Begin filling from row 2 now


# ----------------------------
# UI HEADER (centered)
# ----------------------------
st.markdown("<h1 style='text-align:center; margin-bottom:0.2rem;'>eBay-US Masterfile Filler</h1>", unsafe_allow_html=True)
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
    **Column 1 is TEMPLATE** and **Column 2 is RAW**.
    Flexible header detection is supported; if ambiguous, we assume the first two columns
    are (template, raw) in that order.
    """
    if df_map is None or df_map.empty:
        raise ValueError("Mapping file appears empty. Provide at least two columns: TEMPLATE then RAW.")

    original_cols = list(df_map.columns)
    df_map.columns = [normalize(c) for c in df_map.columns]
    cols = list(df_map.columns)

    template_keys = ("template", "template_header", "masterfile", "master", "target", "masterfile_column", "ebay", "ebay_column", "ebay_header")
    raw_keys = ("raw", "raw_header", "source", "raw_column", "input", "source_column", "rawsheet_column")

    template_col = next((c for c in cols if any(k in c for k in template_keys)), None)
    raw_col = next((c for c in cols if any(k in c for k in raw_keys)), None)

    if template_col is None or raw_col is None:
        # Fallback strict order: first=template, second=raw
        if len(cols) < 2:
            raise ValueError(
                "Mapping file must have at least two columns (TEMPLATE first, RAW second). "
                f"Found columns: {original_cols}"
            )
        template_col, raw_col = cols[0], cols[1]

    mapping = {}
    for _, row in df_map.iterrows():
        tmpl_h = str(row[template_col]).strip() if pd.notna(row[template_col]) else ""
        raw_h = str(row[raw_col]).strip() if pd.notna(row[raw_col]) else ""
        if tmpl_h and raw_h and normalize(tmpl_h) != "nan" and normalize(raw_h) != "nan":
            # We want {template -> raw}
            mapping[tmpl_h] = raw_h
    if not mapping:
        raise ValueError("No valid mappings detected in mapping file (TEMPLATE first, RAW second).")
    return mapping

def get_template_headers(ws: Worksheet) -> Tuple[Dict[int, str], Dict[str, int], int]:
    """
    Read row-1 headers from the first sheet (mapping uses these).
    Returns:
      - by_col: {col_idx -> header_text}
      - by_name: {normalized_header_text -> col_idx}
      - max_col: ws.max_column (for convenience)
    """
    max_col = ws.max_column
    by_col = {}
    by_name = {}
    for c in range(1, max_col + 1):
        v = ws.cell(row=HEADERS_ROW, column=c).value
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
    - Mapping is from row-1 headers (unchanged).
    - Data filling begins from ROW 2 (row 2 now serves as the defaults row).
    - Defaults in row 2 are preserved and copied down to each SKU row.
    - All other sheets remain untouched.
    Returns: (num_rows_written, skipped_columns_due_to_defaults)
    """
    ws: Worksheet = wb.worksheets[0]  # first sheet only

    # Headers from row-1
    by_col, by_name, max_col = get_template_headers(ws)

    # Identify protected columns (i.e., default exists in row 2 and must not be overwritten)
    protected_cols = set()
    defaults_present_cols = set()
    for c in range(1, max_col + 1):
        if nonempty(ws.cell(row=DEFAULTS_SRC_ROW, column=c)) or is_cell_formula(ws.cell(row=DEFAULTS_SRC_ROW, column=c)):
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

    # Track which mapped template columns we skip because row-2 has defaults
    skipped_due_to_defaults = set()

    # Write rows starting at row=2
    for i in range(n_rows):
        target_row = START_WRITE_ROW + i

        # 1) First, copy defaults from row 2 to target_row for ALL columns that have defaults
        if target_row > DEFAULTS_SRC_ROW:
            for c in defaults_present_cols:
                copy_default_cell_to_row(ws, src_row=DEFAULTS_SRC_ROW, dst_row=target_row, col_idx=c)

        # 2) Then write mapped values for this row into non-protected columns only
        for col_idx, raw_col in template_to_raw_pairs:
            if col_idx in protected_cols:
                # do not overwrite defaults
                skipped_due_to_defaults.add(ws.cell(row=HEADERS_ROW, column=col_idx).value)
                continue
            cell = ws.cell(row=target_row, column=col_idx)

            # Copy style from row 2 to keep formatting consistent even in non-default columns
            copy_style(ws.cell(row=DEFAULTS_SRC_ROW, column=col_idx), cell)

            val = data_df.iloc[i][raw_col]
            if pd.isna(val):
                val = None
            cell.value = val

    # Prepare skipped columns (names) for UI display
    skipped_names = sorted(set(name for name in skipped_due_to_defaults if name))
    return n_rows, skipped_names


# ----------------------------
# INPUT SECTIONS (3 required)
# ----------------------------
st.subheader("1) Upload Template")
template_file = st.file_uploader("Upload the eBay-US masterfile template", type=ACCEPTED_TEMPLATE_TYPES, key="template_upload")

st.subheader("2) Upload Raw sheet/ PXM Report (CSV/XLSX)")
raw_file = st.file_uploader("Upload Raw sheet", type=ACCEPTED_RAW_TYPES, key="raw_upload")

st.subheader("3) Upload Mapping file")
mapping_file = st.file_uploader("Upload Mapping file", type=ACCEPTED_MAP_TYPES, key="map_upload")

with st.expander("Mapping file guidance", expanded=False):
    st.markdown(
        """
        **Required order:** the **first column is TEMPLATE** and the **second column is RAW**.  
        Flexible names are supported, but order matters when headers are ambiguous.

        **Example CSV:**
        ```
        template,raw
        SKU,Sku
        Title,Product Title
        Price,Price
        Quantity,Qty
        ```
        - Mapping is matched to **row‑1 headers** of the Template.  
        - Data writing begins at **row 2** (row 2 now serves as the defaults row).  
        - Any **defaults in row 2** (e.g., Localized For, Condition, Measurement System, etc) are **preserved** and copied down for each SKU row.  
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
            st.success(f"Filled {num_rows} row(s) into the first sheet starting at row 2.")
            if skipped_cols:
                st.info(
                    "These mapped template columns were **not** overwritten because row‑2 contains defaults: "
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
