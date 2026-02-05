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
    "Locked, board-ready spreadsheet view for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# FILE CONFIG
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
# COLUMN NAME OVERRIDES (APPROVED)
# =========================================================
COLUMN_NAME_OVERRIDES = {
    "Gemology": {
        "Unnamed: 1": "Instrument Name",
        "Unnamed: 2": "Type",
        "Unnamed: 3": "Specification",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Unit Price (In Lakhs)",
        "Unnamed: 6": "Total in Lakhs",
        "Unnamed: 7": "Remarks"
    },
    "Manufacturing": {
        "Unnamed: 1": "Particulars",
        "Unnamed: 2": "Department",
        "Unnamed: 3": "Details",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Cost per Item (In Lakhs)",
        "Unnamed: 6": "According to Capacity",
        "Unnamed: 7": "Cost-to-Company (In Lakhs)",
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
# HIGHLIGHT RULES (APPROVED)
# =========================================================
GEMOLOGY_CONFIG_ROWS = [
    "number of students",
    "class requirements",
    "singular instruments",
    "instruments per student",
    "shared instruments",
    "room preparation",
    "student per table",
    "faculty required"
]

SECTION_HEADER_KEYWORDS = [
    "instrument name",
    "specification",
    "quantity",
    "unit price",
    "total",
    "remarks"
]

# =========================================================
# SAFE LOADER
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"File not found: {path}")
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

# =========================================================
# LOAD DATA
# =========================================================
cfg = DEPARTMENT_CONFIG[department]
df = load_excel(cfg["file"], cfg["sheet"])

# Apply column name overrides
df.rename(columns=COLUMN_NAME_OVERRIDES.get(department, {}), inplace=True)

# =========================================================
# AG GRID WITH STATIC HIGHLIGHTS
# =========================================================
highlight_js = JsCode("""
function(params) {
    const rowText = Object.values(params.data)
        .join(" ")
        .toLowerCase();

    const firstCell = Object.values(params.data)[0].toLowerCase();

    if (
        ["instrument name", "specification", "quantity", "unit price", "total"]
            .every(k => rowText.includes(k))
    ) {
        return { backgroundColor: "#fff2cc", fontWeight: "bold" };
    }

    if (
        ["total", "grand total"].some(k => rowText.includes(k))
    ) {
        return { backgroundColor: "#fff4b8", fontWeight: "bold" };
    }

    if (
        params.context.department === "Gemology" &&
        ["number of students", "faculty required", "shared instruments"]
            .some(k => firstCell.includes(k))
    ) {
        return { backgroundColor: "#e6f4ea", fontWeight: "bold" };
    }

    return null;
}
""")

gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(
    wrapText=True,
    autoHeight=True,
    resizable=True,
    cellStyle=highlight_js
)

gb.configure_grid_options(
    context={"department": department},
    rowHeight=38
)

# =========================================================
# RENDER
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption("Excel-like view with approved highlights and clean headers.")

AgGrid(
    df,
    gridOptions=gb.build(),
    height=650,
    theme="alpine",
    allow_unsafe_jscode=True
)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption("© Board Excel Intelligence Platform — Locked Rendering Layer")
