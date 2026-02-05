import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder, JsCode

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
# SAFE LOAD
# =========================================================
def load_excel(file, sheet):
    if not os.path.exists(file):
        st.error(f"Required file missing: {file}")
        st.stop()
    return pd.read_excel(file, sheet_name=sheet)

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

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Highlighting (Optional)")

enable_highlight = st.sidebar.checkbox("Enable highlighting", False)
highlight_color = st.sidebar.color_picker("Highlight color", "#FFF3A0")
highlight_row = st.sidebar.number_input(
    "Highlight row (1-based, optional)",
    min_value=0,
    step=1
)

# =========================================================
# LOAD + CLEAN DATA
# =========================================================
config = DEPARTMENT_CONFIG[department]
df = load_excel(config["file"], config["sheet"])
df = df.fillna("")

# Apply column renaming
df = df.rename(columns=COLUMN_NAME_OVERRIDES.get(department, {}))

# =========================================================
# AGGRID CONFIG — TRUE WRAP
# =========================================================
gb = GridOptionsBuilder.from_dataframe(df)

# Default column behavior
gb.configure_default_column(
    wrapText=True,
    autoHeight=True,
    resizable=True,
    filter=True,
    sortable=True,
    cellStyle={
        "white-space": "normal",
        "line-height": "1.4",
        "word-break": "break-word"
    }
)

# Optional row highlighting
if enable_highlight and highlight_row > 0:
    js = JsCode(f"""
        function(params) {{
            if (params.node.rowIndex === {highlight_row - 1}) {{
                return {{
                    backgroundColor: '{highlight_color}',
                    fontWeight: 'bold'
                }}
            }}
        }}
    """)
    gb.configure_grid_options(getRowStyle=js)

grid_options = gb.build()

# =========================================================
# RENDER
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption("Excel-like, read-only view with wrapped text and full visibility.")

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
st.caption(
    "© Board Excel Intelligence Platform — Locked Academic Expansion Dashboard"
)
