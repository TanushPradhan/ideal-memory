import streamlit as st
import pandas as pd
import os
import streamlit.components.v1 as components

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
        "sheet": 0
    },
    "Manufacturing": {
        "file": "IIGJ - Academic Expansion Plan - Manufacturing Formatted.xlsx",
        "sheet": 0
    },
    "CAD": {
        "file": "IIGJ Mumbai - Academic expansion Plan_CAD Formatted.xlsx",
        "sheet": 0
    }
}

# =========================================================
# PRE-HIGHLIGHT RULES
# =========================================================
TOTAL_KEYWORDS = [
    "total",
    "grand total",
    "total (approx.)",
    "all instruments cost",
    "per student instruments",
    "shared resources"
]

SECTION_HEADER_KEYWORDS = [
    "instrument name",
    "type",
    "specification",
    "quantity",
    "unit price",
    "total in lakhs",
    "remarks"
]

TOTAL_HIGHLIGHT_COLOR = "#fff4b8"
SECTION_HEADER_COLOR = "#fff2cc"

# =========================================================
# SAFE EXCEL LOADER
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"🚫 Required file not found:\n{path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet, header=None).fillna("")

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
config = DEPARTMENT_CONFIG[department]
df = load_excel(config["file"], config["sheet"])

# =========================================================
# HTML TABLE RENDERER
# =========================================================
def render_html_table(df):
    html = f"""
    <style>
        .excel-container {{
            max-height: 75vh;
            overflow: auto;
            border: 1px solid #bfbfbf;
            background-color: #ffffff;
        }}

        table {{
            border-collapse: collapse;
            width: 100%;
            font-size: 13px;
            font-family: Arial, Helvetica, sans-serif;
            color: #000000;
            background-color: #ffffff;
        }}

        th, td {{
            border: 1px solid #c0c0c0;
            padding: 6px 8px;
            vertical-align: top;
            text-align: left;
            white-space: pre-wrap;
            word-wrap: break-word;
        }}

        tr.total-row td {{
            background-color: {TOTAL_HIGHLIGHT_COLOR};
            font-weight: 600;
        }}

        tr.section-header td {{
            background-color: {SECTION_HEADER_COLOR};
            font-weight: 600;
        }}

        tr:nth-child(even):not(.total-row):not(.section-header) td {{
            background-color: #fafafa;
        }}
    </style>

    <div class="excel-container">
    <table>
    """

    for _, row in df.iterrows():
        row_text = " ".join(str(cell) for cell in row).lower()

        is_total = any(k in row_text for k in TOTAL_KEYWORDS)
        is_section_header = sum(
            1 for k in SECTION_HEADER_KEYWORDS if k in row_text
        ) >= 4  # strong match threshold

        row_class = ""
        if is_section_header:
            row_class = "section-header"
        elif is_total:
            row_class = "total-row"

        html += f'<tr class="{row_class}">'

        for cell in row:
            html += f"<td>{cell}</td>"

        html += "</tr>"

    html += "</table></div>"
    return html

# =========================================================
# DISPLAY
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption(
    "Excel-like, read-only view with wrapped text, gridlines, "
    "section headers, and highlighted totals."
)

components.html(
    render_html_table(df),
    height=800,
    scrolling=True
)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Spreadsheet Rendering Layer"
)
