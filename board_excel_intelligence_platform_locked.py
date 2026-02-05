import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder

# =========================================================
# APP CONFIG
# =========================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

st.title("📊 Board Excel Intelligence Platform")
st.caption(
    "Locked, board-ready spreadsheet view for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# DEPARTMENT CONFIG
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
# SAFE LOADER
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"File not found: {path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet)

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
df = df.fillna("")

# =========================================================
# AG GRID CONFIG (THIS IS THE MAGIC)
# =========================================================
def render_excel_grid(dataframe):
    gb = GridOptionsBuilder.from_dataframe(dataframe)

    gb.configure_default_column(
        resizable=True,
        sortable=False,
        filter=True,
        wrapText=True,
        autoHeight=True,
        cellStyle={
            "white-space": "normal",
            "line-height": "1.4",
            "border": "1px solid #d0d0d0"
        }
    )

    gb.configure_grid_options(
        domLayout="normal",
        suppressRowHoverHighlight=False,
        rowHeight=38
    )

    AgGrid(
        dataframe,
        gridOptions=gb.build(),
        fit_columns_on_grid_load=False,
        height=650,
        theme="alpine",
        allow_unsafe_jscode=True
    )

# =========================================================
# EXECUTIVE INSIGHTS
# =========================================================
def render_insights(dept):
    st.markdown("## 🧠 Executive Insights")
    st.markdown("---")

    if dept == "Gemology":
        st.markdown("""
- **Student capacity:** ~60–65 students
- **Shared instruments** reduce per-student cost
- **High compliance focus** (room sealing, lighting, lab standards)
- **Capital efficient** with long asset life
""")

    elif dept == "Manufacturing":
        st.markdown("""
- **Total estimated capex:** ₹166+ Lakhs
- Heavy machinery dominates cost structure
- Shared capacity logic improves utilization
- High operational dependency
""")

    elif dept == "CAD":
        st.markdown("""
- Lowest infrastructure cost
- Software-driven recurring expenses
- Highest scalability potential
""")

# =========================================================
# RENDER
# =========================================================
if view_mode == "Executive View":
    render_insights(department)
    st.markdown("---")

st.subheader(f"📄 Spreadsheet View — {department}")
st.caption("Excel-like, read-only view with wrapped text and full gridlines.")

render_excel_grid(df)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption("© Board Excel Intelligence Platform — Spreadsheet Rendering Layer")
