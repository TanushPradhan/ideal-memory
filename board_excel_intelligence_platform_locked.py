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
# HIGHLIGHT RULES (LOCKED)
# =========================================================
TOTAL_KEYWORDS = ["total", "grand total", "total (approx.)"]

SECTION_HEADER_KEYWORDS = [
    "instrument name", "type", "specification",
    "quantity", "unit price", "total in lakhs", "remarks"
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

TOTAL_HIGHLIGHT_COLOR = "#fff4b8"
SECTION_HEADER_COLOR = "#fff2cc"
CONFIG_HIGHLIGHT_COLOR = "#e6f4ea"

# =========================================================
# SAFE EXCEL LOADER
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"🚫 Required file not found:\n{path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet, header=None).fillna("")

def is_number(val):
    try:
        float(val)
        return True
    except:
        return False

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
# HTML TABLE RENDERER (LOCKED STRUCTURE)
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
            color: #000000;
        }}

        th, td {{
            border: 1px solid #c0c0c0;
            padding: 6px 8px;
            vertical-align: top;
            white-space: pre-wrap;
            word-wrap: break-word;
        }}

        td.numeric-cell {{
            text-align: center;
            white-space: nowrap;
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
        row_text = " ".join(str(cell) for cell in row).lower()
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
            cell_class = "numeric-cell" if is_number(cell) else ""
            html += f'<td class="{cell_class}">{cell}</td>'
        html += "</tr>"

    html += "</table></div>"
    return html

# =========================================================
# DISPLAY TABLE
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption(
    "Excel-like, read-only view with wrapped text, gridlines, "
    "section headers, configuration blocks, totals, and centered numerics."
)

components.html(
    render_html_table(df, department),
    height=800,
    scrolling=True
)

# =========================================================
# EXECUTIVE INSIGHTS (STATIC)
# =========================================================
st.markdown("---")
st.subheader("🧠 Executive Insights")

if department == "Gemology":
    st.markdown("""
• Structured for ~65 students with spare capacity  
• Hybrid per-student + shared precision instruments  
• Security-first lab design (sealed rooms)  
• Faculty ratio aligned to accreditation norms  
""")

elif department == "Manufacturing":
    st.markdown("""
• Capital-intensive shared machinery model  
• Equipment sized for 60-student throughput  
• Consumables pooled to reduce redundancy  
• Clear separation of per-student vs shared costs  
""")

elif department == "CAD":
    st.markdown("""
• Digital-first infrastructure  
• Lower physical capex than labs/workshops  
• Linear scalability with student growth  
• Shared backend compute resources  
""")

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Spreadsheet Rendering Layer + Executive Intelligence Layer"
)
