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
# HIGHLIGHT RULES (APPROVED EARLIER)
# =========================================================
TOTAL_KEYWORDS = ["total", "grand total", "total (approx.)"]

SECTION_HEADER_KEYWORDS = [
    "instrument name", "type", "specification",
    "quantity", "unit price", "total", "remarks"
]

GEMOLOGY_CONFIG_ROWS = [
    "number of students",
    "class requirements with spare",
    "singular instruments",
    "instruments per student",
    "shared instruments required",
    "room preparation",
    "student per table",
    "faculty required"
]

TOTAL_HIGHLIGHT_COLOR = "#fff4b8"     # light yellow
SECTION_HEADER_COLOR = "#e8f2ff"      # light blue
CONFIG_HIGHLIGHT_COLOR = "#e6f4ea"     # light green

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
# HTML TABLE RENDERER (RESTORED + FINAL)
# =========================================================
def render_html_table(df, department):
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
            color: #000;
        }}

        td {{
            border: 1px solid #c0c0c0;
            padding: 6px 8px;
            vertical-align: top;
            white-space: pre-wrap;
            word-wrap: break-word;
        }}

        /* Alignment */
        td:nth-child(1),
        td:nth-child(2),
        td:nth-child(9) {{
            text-align: left;
        }}

        td:nth-child(3),
        td:nth-child(4),
        td:nth-child(5),
        td:nth-child(6),
        td:nth-child(7),
        td:nth-child(8) {{
            text-align: center;
        }}

        tr.total-row td {{
            background-color: {TOTAL_HIGHLIGHT_COLOR};
            font-weight: 600;
        }}

        tr.section-header td {{
            background-color: {SECTION_HEADER_COLOR};
            font-weight: 600;
        }}

        tr.config-row td {{
            background-color: {CONFIG_HIGHLIGHT_COLOR};
            font-weight: 600;
        }}

        tr:nth-child(even):not(.total-row):not(.section-header):not(.config-row) td {{
            background-color: #fafafa;
        }}
    </style>

    <div class="excel-container">
    <table>
    """

    for _, row in df.iterrows():
        row_text = " ".join(str(c) for c in row).lower()
        first_cell = str(row.iloc[0]).lower().strip()

        is_total = any(k in row_text for k in TOTAL_KEYWORDS)
        is_section_header = sum(1 for k in SECTION_HEADER_KEYWORDS if k in row_text) >= 4
        is_config = (
            department == "Gemology"
            and any(first_cell.startswith(k) for k in GEMOLOGY_CONFIG_ROWS)
        )

        row_class = ""
        if is_config:
            row_class = "config-row"
        elif is_section_header:
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
    "section headers, configuration blocks, and totals."
)

components.html(
    render_html_table(df, department),
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
