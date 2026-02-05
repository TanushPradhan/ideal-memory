import streamlit as st
import pandas as pd
import os

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
# GLOBAL TABLE STYLING (TEXT WRAP + ALIGNMENT)
# =========================================================
st.markdown(
    """
    <style>
    thead th {
        white-space: normal !important;
        word-wrap: break-word !important;
        text-align: center !important;
    }
    tbody td {
        white-space: normal !important;
        word-wrap: break-word !important;
        vertical-align: top !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================================================
# LOCKED DEPARTMENT CONFIG
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
        "sheet": "Formated_Budget"
    }
}

# =========================================================
# PROFESSIONAL COLUMN NAME OVERRIDES (BOARD SAFE)
# =========================================================
COLUMN_NAME_OVERRIDES = {
    "Gemology": {
        "Unnamed: 1": "Instrument Name",
        "Unnamed: 2": "Specification",
        "Unnamed: 3": "Students",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Unit Price (in Lakhs)",
        "Unnamed: 6": "Total (in lakhs)",
        "Unnamed: 7": "Remarks"
    },
    "Manufacturing": {
        "Unnamed: 1": "Particulars",
        "Unnamed: 2": "Department",
        "Unnamed: 3": "Details",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Cost per Item (In Lakhs)",
        "Unnamed: 6": "According to the capacity requirement for 60 students.",
        "Unnamed: 7": "Cost-to-Company in Lakhs (Equipment according to the capacity requirement of 60)",
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
def load_excel_safely(path, sheet):
    if not os.path.exists(path):
        st.error(f"🚫 Required board file not found:\n\n{path}")
        st.stop()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        st.error(f"🚫 Unable to read Excel sheet.\n\n{e}")
        st.stop()

# =========================================================
# CLEAN COLUMN HEADERS
# =========================================================
def clean_column_headers(df, department):
    overrides = COLUMN_NAME_OVERRIDES.get(department, {})
    df.columns = [
        overrides.get(col, col) if isinstance(col, str) else col
        for col in df.columns
    ]
    return df

# =========================================================
# SIDEBAR – NAVIGATION
# =========================================================
st.sidebar.header("📁 Navigation")

selected_department = st.sidebar.selectbox(
    "Select Department",
    list(DEPARTMENT_CONFIG.keys())
)

view_mode = st.sidebar.radio(
    "View Mode",
    ["Interactive Spreadsheet", "Executive View"]
)

st.sidebar.markdown("---")
st.sidebar.subheader("🎨 Highlighting (Optional)")

enable_highlight = st.sidebar.checkbox("Enable highlighting")

highlight_color = st.sidebar.color_picker(
    "Highlight color",
    "#FFF3A0",
    disabled=not enable_highlight
)

highlight_row = st.sidebar.number_input(
    "Highlight row (1-based, optional)",
    min_value=0,
    step=1,
    disabled=not enable_highlight
)

# =========================================================
# LOAD DATA
# =========================================================
config = DEPARTMENT_CONFIG[selected_department]
df = load_excel_safely(config["file"], config["sheet"])
df = clean_column_headers(df, selected_department)
df_display = df.fillna("")

# =========================================================
# STYLING FUNCTION (NO AUTO HIGHLIGHTING)
# =========================================================
def style_dataframe(df):
    styled = df.style.set_properties(
        **{
            "text-align": "left",
            "white-space": "normal",
            "word-wrap": "break-word"
        }
    )

    if enable_highlight and highlight_row > 0 and highlight_row <= len(df):
        styled = styled.apply(
            lambda x: [
                f"background-color: {highlight_color}; font-weight: bold;"
                if i == highlight_row - 1 else ""
                for i in range(len(df))
            ],
            axis=0
        )

    return styled

# =========================================================
# INTERACTIVE VIEW
# =========================================================
if view_mode == "Interactive Spreadsheet":
    st.subheader(f"📄 Spreadsheet View — {selected_department}")
    st.caption("Excel-like, read-only view. No assumptions or calculations.")

    st.dataframe(
        df_display,
        use_container_width=True,
        height=650
    )

# =========================================================
# EXECUTIVE VIEW
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {selected_department}")
    st.caption("Board-friendly structured view with clarity and completeness.")

    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()

    if numeric_cols:
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Rows", len(df))
        col2.metric("Numeric Columns", len(numeric_cols))
        col3.metric("Total Numeric Sum", f"{df[numeric_cols].sum().sum():,.2f}")

    st.markdown("---")

    st.dataframe(
        style_dataframe(df_display),
        use_container_width=True,
        height=650
    )

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Locked Academic Expansion Dashboard"
)

