import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# =====================================================
# PAGE CONFIG
# =====================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

st.title("📊 Board Excel Intelligence Platform")
st.caption(
    "Locked, board-ready dashboard for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =====================================================
# DATA SOURCE CONFIG (LOCKED – DO NOT GENERALISE)
# =====================================================
BASE_PATH = Path(__file__).parent

DATA_SOURCES = {
    "Gemology": {
        "file": BASE_PATH / "IIGJ_Gemology_Formatted.xlsx",
        "sheet": "Budget",
        "numeric_columns": [
            "No of Students",
            "Unit Cost (₹ Lakhs)",
            "Total Cost (₹ Lakhs)"
        ],
        "primary_text_column": "Particulars"
    },
    "Manufacturing": {
        "file": BASE_PATH / "IIGJ_Manufacturing_Formatted.xlsx",
        "sheet": "Budget",
        "numeric_columns": [
            "Quantity",
            "Unit Cost (₹ Lakhs)",
            "Total Cost (₹ Lakhs)"
        ],
        "primary_text_column": "Particulars"
    },
    "CAD": {
        "file": BASE_PATH / "IIGJ_CAD_Formatted.xlsx",
        "sheet": "Budget",
        "numeric_columns": [
            "No of Students",
            "Unit Cost (₹ Lakhs)",
            "Total Cost (₹ Lakhs)"
        ],
        "primary_text_column": "Particulars"
    }
}

# =====================================================
# SIDEBAR – NAVIGATION
# =====================================================
st.sidebar.header("📁 Navigation")

department = st.sidebar.selectbox(
    "Select Department",
    list(DATA_SOURCES.keys())
)

view_mode = st.sidebar.radio(
    "View Mode",
    ["Interactive Spreadsheet", "Executive View"]
)

st.sidebar.header("🎨 Highlighting")

highlight_color = st.sidebar.color_picker(
    "Highlight numeric columns",
    "#FFF3B0"
)

highlight_row = st.sidebar.number_input(
    "Highlight row (1-based, optional)",
    min_value=0,
    step=1
)

# =====================================================
# LOAD DATA (DETERMINISTIC & SAFE)
# =====================================================
config = DATA_SOURCES[department]

df = pd.read_excel(
    config["file"],
    sheet_name=config["sheet"]
)

# Keep blanks blank (no NaN display)
df = df.replace({np.nan: ""})

numeric_columns = config["numeric_columns"]
primary_text_column = config["primary_text_column"]

# =====================================================
# AGGRID CONFIGURATION
# =====================================================
gb = GridOptionsBuilder.from_dataframe(df)

for col in df.columns:
    is_numeric = col in numeric_columns
    is_primary = col == primary_text_column

    cell_style = {
        "textAlign": "center" if is_numeric else "left",
        "whiteSpace": "normal",
        "lineHeight": "1.4",
        "borderRight": "1px solid #3a3a3a",
        "borderBottom": "1px solid #2a2a2a",
    }

    if is_primary:
        cell_style["fontWeight"] = "600"

    if is_numeric:
        cell_style["backgroundColor"] = highlight_color

    gb.configure_column(
        col,
        wrapText=True,
        autoHeight=True,
        cellStyle=cell_style
    )

# =====================================================
# ROW HIGHLIGHTING (SAFE & OPTIONAL)
# =====================================================
if highlight_row > 0 and highlight_row <= len(df):
    df["_row_flag"] = ""
    df.loc[highlight_row - 1, "_row_flag"] = "highlight"

    gb.configure_column("_row_flag", hide=True)

    gb.configure_grid_options(
        getRowStyle={
            "styleConditions": [
                {
                    "condition": "params.data._row_flag === 'highlight'",
                    "style": {
                        "backgroundColor": "#E3F2FD",
                        "fontWeight": "600"
                    }
                }
            ]
        }
    )

gb.configure_grid_options(
    suppressColumnVirtualisation=True,
    alwaysShowHorizontalScroll=True
)

grid_options = gb.build()

# =====================================================
# INTERACTIVE SPREADSHEET VIEW
# =====================================================
if view_mode == "Interactive Spreadsheet":
    st.subheader(f"📄 {department} – Spreadsheet View")

    AgGrid(
        df,
        gridOptions=grid_options,
        height=560,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        theme="streamlit"
    )

# =====================================================
# EXECUTIVE VIEW
# =====================================================
else:
    st.subheader(f"📊 {department} – Executive Summary")

    numeric_df = df[numeric_columns].apply(
        pd.to_numeric, errors="coerce"
    )

    total_cost = numeric_df.sum().sum()
    highest_cost = numeric_df.max().max()

    col1, col2 = st.columns(2)
    col1.metric(
        "Total Investment (₹ Lakhs)",
        f"{round(total_cost, 2)}"
    )
    col2.metric(
        "Highest Line Item (₹ Lakhs)",
        f"{round(highest_cost, 2)}"
    )

    st.markdown("---")

    AgGrid(
        df,
        gridOptions=grid_options,
        height=560,
        update_mode=GridUpdateMode.NO_UPDATE,
        fit_columns_on_grid_load=False,
        theme="streamlit"
    )
