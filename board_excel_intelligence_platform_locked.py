import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder

# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

# =========================================================
# GLOBAL CSS (GRID LINES + WRAP TEXT)
# =========================================================
st.markdown(
    """
    <style>
    .ag-theme-balham-dark {
        --ag-row-border-color: #3a3a3a;
        --ag-cell-horizontal-border: solid #3a3a3a;
        --ag-header-column-separator-display: block;
        --ag-header-column-separator-color: #3a3a3a;
    }

    .ag-theme-balham-dark .ag-cell {
        white-space: normal !important;
        line-height: 1.4 !important;
        border-right: 1px solid #3a3a3a !important;
        border-bottom: 1px solid #3a3a3a !important;
    }

    .ag-theme-balham-dark .ag-header-cell {
        white-space: normal !important;
        line-height: 1.3 !important;
        font-weight: 600;
        border-right: 1px solid #3a3a3a !important;
        border-bottom: 1px solid #3a3a3a !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# TITLE
# =========================================================
st.title("📊 Board Excel Intelligence Platform")
st.caption(
    "Locked, board-ready dashboard for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# FILE CONFIG (LOCKED)
# =========================================================
DEPARTMENT_CONFIG = {
    "Gemology": {
        "file": "IIGJ Mumbai - Academic Expansion Plan - Gemology Formatted.xlsx",
        "sheet": "Department of Gemmology_Format"
    },
    "Manufacturing": {
        "file": "IIGJ - Academic Expansion Plan - Manufacturing Formatted.xlsx",
        "sheet": "Formated_Budget"
    },
    "CAD": {
        "file": "IIGJ Mumbai - Academic expansion Plan_CAD Formatted.xlsx",
        "sheet": "Formatted_Budget"
    }
}

# =========================================================
# COLUMN NAME OVERRIDES (BOARD SAFE)
# =========================================================
COLUMN_NAME_OVERRIDES = {
    "Gemology": {
        "Unnamed: 1": "Instrument Name",
        "Unnamed: 2": "Specification",
        "Unnamed: 3": "Students",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Unit Price (in Lakhs)",
        "Unnamed: 6": "Total (in Lakhs)",
        "Unnamed: 7": "Remarks"
    },
    "Manufacturing": {
        "Unnamed: 1": "Particulars",
        "Unnamed: 2": "Department",
        "Unnamed: 3": "Details",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Cost per Item (in Lakhs)",
        "Unnamed: 6": "According to Capacity Requirement (60 Students)",
        "Unnamed: 7": "Cost-to-Company (in Lakhs)",
        "Unnamed: 8": "Remarks"
    },
    "CAD": {
        "Unnamed: 1": "Particulars",
        "Unnamed: 2": "Specification",
        "Unnamed: 3": "Quantity",
        "Unnamed: 4": "Cost per Item",
        "Unnamed: 5": "Total Cost",
        "Unnamed: 6": "Remarks"
    }
}

# =========================================================
# SAFE EXCEL LOADER
# =========================================================
def load_excel(file_path, sheet):
    if not os.path.exists(file_path):
        st.error(f"Required file not found: {file_path}")
        st.stop()
    return pd.read_excel(file_path, sheet_name=sheet)

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("📁 Navigation")

department = st.sidebar.selectbox(
    "Select Department",
    list(DEPARTMENT_CONFIG.keys())
)

view_mode = st.sidebar.radio(
    "View Mode",
    ["Interactive Spreadsheet", "Executive View"]
)

# =========================================================
# LOAD DATA
# =========================================================
config = DEPARTMENT_CONFIG[department]
df = load_excel(config["file"], config["sheet"])

# Rename columns safely
df = df.rename(columns=COLUMN_NAME_OVERRIDES.get(department, {}))
df = df.fillna("")

# =========================================================
# GRID OPTIONS (WRAP + GRIDLINES)
# =========================================================
gb = GridOptionsBuilder.from_dataframe(df)

gb.configure_default_column(
    wrapText=True,
    autoHeight=True,
    resizable=True,
    sortable=False,
    filter=True
)

gb.configure_grid_options(
    domLayout="normal",
    suppressRowHoverHighlight=False
)

grid_options = gb.build()

# =========================================================
# INTERACTIVE VIEW
# =========================================================
if view_mode == "Interactive Spreadsheet":
    st.subheader(f"📄 Spreadsheet View — {department}")
    st.caption("Excel-like, read-only view with wrapped text.")

    AgGrid(
        df,
        gridOptions=grid_options,
        theme="balham-dark",
        height=700,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True
    )

# =========================================================
# EXECUTIVE VIEW
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {department}")
    st.caption("Board-ready view with wrapped text and clear grid lines.")

    numeric_cols = df.select_dtypes(include=["number"]).columns
    if len(numeric_cols) > 0:
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Rows", len(df))
        c2.metric("Numeric Columns", len(numeric_cols))
        c3.metric("Total Numeric Sum", f"{df[numeric_cols].sum().sum():,.2f}")

    st.markdown("---")

    AgGrid(
        df,
        gridOptions=grid_options,
        theme="balham-dark",
        height=700,
        fit_columns_on_grid_load=False,
        allow_unsafe_jscode=True
    )

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption("© Board Excel Intelligence Platform — Locked Academic Expansion Dashboard")
