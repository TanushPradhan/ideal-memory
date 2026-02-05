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
# GLOBAL CSS (TEXT WRAP + CLEAN TABLE)
# =========================================================
st.markdown("""
<style>
    td {
        white-space: normal !important;
        word-wrap: break-word;
        vertical-align: top;
    }
    th {
        white-space: normal !important;
        word-wrap: break-word;
        text-align: center;
    }
</style>
""", unsafe_allow_html=True)

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
# SAFE EXCEL LOADER
# =========================================================
def load_excel_safely(path, sheet):
    if not os.path.exists(path):
        st.error(f"🚫 Required board file not found:\n\n{path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet)

# =========================================================
# SIDEBAR
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

enable_highlight = st.sidebar.checkbox(
    "Enable highlighting",
    value=False
)

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
df_display = df.fillna("")

# =========================================================
# STYLING FUNCTION (WRAP + ALIGNMENT)
# =========================================================
def style_dataframe(df):
    def style_cell(val):
        if isinstance(val, (int, float)):
            base = "text-align: center;"
            if enable_highlight:
                return base + f" background-color: {highlight_color};"
            return base
        return "text-align: left;"

    styled = (
        df.style
        .applymap(style_cell)
        .set_properties(**{
            "white-space": "normal",
            "word-wrap": "break-word",
            "vertical-align": "top"
        })
    )

    if enable_highlight and 0 < highlight_row <= len(df):
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
    st.caption("Excel-like, read-only view. No calculations or assumptions.")

    st.dataframe(
        df_display,
        use_container_width=True,
        height=650
    )

# =========================================================
# EXECUTIVE VIEW (WRAPPED, BOARD-READY)
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {selected_department}")
    st.caption("Board-friendly structured view with wrapped text and alignment.")

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
