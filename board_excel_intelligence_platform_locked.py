import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# =========================================================
# APP CONFIG
# =========================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

st.title("📊 Board Excel Intelligence Platform")
st.caption(
    "Locked, board-ready dashboard for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# LOCKED FILE CONFIG
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
# PROFESSIONAL COLUMN NAME OVERRIDES
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
        "Unnamed: 6": "According to Capacity (60 Students)",
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
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"Required file missing: {path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet).fillna("")

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
    ["Analysis View", "Executive View"]
)

# =========================================================
# LOAD & PREP DATA
# =========================================================
config = DEPARTMENT_CONFIG[department]
df = load_excel(config["file"], config["sheet"])

# Rename columns safely
df = df.rename(columns=COLUMN_NAME_OVERRIDES.get(department, {}))

# =========================================================
# AGGRID BUILDER (CORE)
# =========================================================
def render_grid(df, compact=False):
    gb = GridOptionsBuilder.from_dataframe(df)

    gb.configure_default_column(
        wrapText=True,
        autoHeight=True,
        resizable=True,
        sortable=True,
        filter=True,
        cellStyle={
            "borderRight": "1px solid #d0d0d0",
            "borderBottom": "1px solid #d0d0d0",
            "white-space": "normal",
            "line-height": "1.4"
        }
    )

    gb.configure_grid_options(
        domLayout="normal",
        suppressRowHoverHighlight=False
    )

    if compact:
        gb.configure_grid_options(rowHeight=32)
    else:
        gb.configure_grid_options(rowHeight=48)

    AgGrid(
        df,
        gridOptions=gb.build(),
        theme="alpine",
        update_mode=GridUpdateMode.NO_UPDATE,
        allow_unsafe_jscode=True,
        fit_columns_on_grid_load=False,
        height=650
    )

# =========================================================
# ANALYSIS VIEW
# =========================================================
if view_mode == "Analysis View":
    st.subheader(f"📄 Analysis View — {department}")
    st.caption("Analyst-focused view for fast inspection and verification.")

    render_grid(df, compact=True)

# =========================================================
# EXECUTIVE VIEW
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {department}")
    st.caption(
        "Board-ready presentation with wrapped text, clear column separators "
        "and comfortable spacing."
    )

    # Executive metrics (safe & optional)
    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if numeric_cols:
        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", len(df))
        c2.metric("Numeric Columns", len(numeric_cols))
        c3.metric("Total (All Numbers)", f"{df[numeric_cols].sum().sum():,.2f}")

    st.markdown("---")

    render_grid(df, compact=False)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Locked Academic Expansion Dashboard"
)
