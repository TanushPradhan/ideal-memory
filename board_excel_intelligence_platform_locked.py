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
        "sheet": "Formatted_Budget"
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
        "Unnamed: 6": "Total (in Lakhs)",
        "Unnamed: 7": "Remarks"
    },
    "Manufacturing": {
        "Unnamed: 1": "Particulars",
        "Unnamed: 2": "Department",
        "Unnamed: 3": "Details",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Cost per Item (in Lakhs)",
        "Unnamed: 6": "Capacity Requirement (60 Students)",
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
def load_excel_safely(path, sheet):
    if not os.path.exists(path):
        st.error(
            f"🚫 Required board file not found:\n\n{path}\n\n"
            "Please contact the administrator."
        )
        st.stop()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        st.error(
            "🚫 Unable to read the configured Excel sheet.\n\n"
            f"Details: {e}"
        )
        st.stop()

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

# =========================================================
# LOAD DATA
# =========================================================
config = DEPARTMENT_CONFIG[selected_department]
df = load_excel_safely(config["file"], config["sheet"])

# Rename columns professionally
df = df.rename(columns=COLUMN_NAME_OVERRIDES.get(selected_department, {}))

# Replace NaN with empty strings for clean display
df_display = df.fillna("")

# =========================================================
# STYLING (NO AUTO HIGHLIGHTING)
# =========================================================
def style_dataframe(df):
    return (
        df.style
        .set_properties(**{
            "white-space": "normal",
            "word-wrap": "break-word",
            "vertical-align": "top",
            "text-align": "left"
        })
        .set_table_styles([
            {
                "selector": "th",
                "props": [
                    ("white-space", "normal"),
                    ("word-wrap", "break-word"),
                    ("text-align", "center"),
                    ("font-weight", "bold")
                ]
            }
        ])
    )

# =========================================================
# INTERACTIVE SPREADSHEET VIEW
# (Fast scrolling, Excel-like)
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
# (Board-safe, wrapped, readable)
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {selected_department}")
    st.caption("Board-friendly structured view with complete visibility.")

    numeric_cols = df.select_dtypes(include=["number"]).columns.tolist()

    if numeric_cols:
        col1, col2, col3 = st.columns(3)
        col1.metric("Total Rows", len(df))
        col2.metric("Numeric Columns", len(numeric_cols))
        col3.metric(
            "Total Numeric Sum",
            f"{df[numeric_cols].sum().sum():,.2f}"
        )

    st.markdown("---")

    # IMPORTANT:
    # st.write + Pandas Styler is REQUIRED for text wrapping
    st.write(style_dataframe(df_display))

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Locked Academic Expansion Dashboard"
)
