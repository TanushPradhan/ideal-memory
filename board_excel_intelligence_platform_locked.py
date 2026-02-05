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
# HIGHLIGHT RULES (LOCKED – APPROVED)
# =========================================================
TOTAL_KEYWORDS = ["total", "grand total", "total (approx.)"]

SECTION_HEADER_KEYWORDS = [
    "instrument name", "type", "specification",
    "quantity", "unit price", "total in lakhs", "remarks"
]

MANUFACTURING_MAIN_HEADER_KEYWORDS = [
    "sl. no.", "particulars", "department", "details",
    "quantity", "cost per items", "according to the capacity",
    "cost-to-company", "remarks"
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
MAIN_HEADER_COLOR = "#e7f3ff"

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
# HTML TABLE RENDERER (FINAL, APPROVED)
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
        td {{
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
        tr.config-row td {{
            background-color: {CONFIG_HIGHLIGHT_COLOR};
            font-weight: 600;
        }}
        tr.main-header td {{
            background-color: {MAIN_HEADER_COLOR};
            font-weight: 700;
        }}
        tr:nth-child(even):not(.total-row):not(.section-header):not(.config-row):not(.main-header) td {{
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
        is_section_header = sum(k in row_text for k in SECTION_HEADER_KEYWORDS) >= 4
        is_main_header = sum(k in row_text for k in MANUFACTURING_MAIN_HEADER_KEYWORDS) >= 6
        is_config = (
            department == "Gemology"
            and any(first_cell.startswith(k) for k in GEMOLOGY_CONFIG_ROWS)
        )

        row_class = ""
        if is_main_header:
            row_class = "main-header"
        elif is_config:
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
# TABLE DISPLAY
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption(
    "Excel-like, read-only view with wrapped text, full gridlines, "
    "section headers, configuration blocks, and totals."
)

components.html(
    render_html_table(df, department),
    height=800,
    scrolling=True
)

# =========================================================
# EXECUTIVE INSIGHTS (TEXT ONLY – STATIC)
# =========================================================
st.markdown("---")
st.subheader("🧑‍💼 Executive Insights")
st.caption("Board-level interpretation of the above data. Figures are indicative and planning-oriented.")

if department == "Gemology":
    st.markdown("""
**Key Observations – Gemology**

• The Gemology program is **student-capacity driven**, with most investments scaling directly with class strength (~65 students).  
• A significant portion of cost is allocated to **per-student resources**, indicating a strong focus on hands-on learning.  
• High-value instruments (microscopes, grading equipment) are **shared assets**, optimizing capital efficiency.  
• Room preparation and infrastructure controls highlight **security and compliance priorities**, critical for gemstone handling.  
• Overall, Gemology demonstrates a **balanced CAPEX model**: high academic rigor with controlled shared investments.
""")

elif department == "Manufacturing":
    st.markdown("""
**Key Observations – Manufacturing**

• Manufacturing shows the **highest absolute capital investment**, driven by heavy machinery and equipment.  
• Most machines are **shared across students**, significantly reducing per-student cost.  
• High-cost items such as casting machines, furnaces, and compressors dominate the budget.  
• Consumables are structured as **shared recurring resources**, improving long-term cost sustainability.  
• The model supports **scalability** — increasing intake does not proportionally increase CAPEX.
""")

elif department == "CAD":
    st.markdown("""
**Key Observations – CAD**

• CAD investment is primarily **software- and system-driven**, rather than physical infrastructure.  
• Costs are largely front-loaded through **licenses and workstations**, with low marginal cost per additional student.  
• Hardware requirements are standardized, allowing predictable budgeting.  
• CAD demonstrates the **lowest operational complexity** among the three departments.  
• This program is **highly scalable** and cost-efficient for expansion.
""")

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Spreadsheet Rendering & Executive Interpretation Layer"
)
