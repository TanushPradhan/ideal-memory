import streamlit as st
import pandas as pd
import os
import streamlit.components.v1 as components
import altair as alt

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
# HIGHLIGHT RULES (APPROVED)
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
# HTML TABLE RENDERER (LOCKED — DO NOT TOUCH)
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
            html += f"<td>{cell}</td>"
        html += "</tr>"

    html += "</table></div>"
    return html

# =========================================================
# DISPLAY TABLE
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
# EXECUTIVE INSIGHTS (STATIC, BOARD-SAFE)
# =========================================================
st.markdown("---")
st.subheader("📊 Executive Insights")
st.caption("Static strategic interpretation for governing council review (₹ in Lakhs).")

INSIGHTS_DATA = {
    "Gemology": {
        "Per-Student Equipment": 48.2,
        "Shared Infrastructure": 72.4,
        "Consumables": 18.6,
        "Infrastructure & Setup": 12.8
    },
    "Manufacturing": {
        "Heavy Machinery": 92.5,
        "Shared Equipment": 48.3,
        "Per-Student Equipment": 21.2,
        "Consumables": 4.0
    },
    "CAD": {
        "Systems & Licenses": 38.0,
        "Per-Student Hardware": 22.0,
        "Shared Infrastructure": 10.0
    }
}

DEPARTMENT_TOTALS = {
    "Gemology": 152.0,
    "Manufacturing": 166.0,
    "CAD": 70.0
}

# -------------------------
# Cost Composition Chart
# -------------------------
composition_df = pd.DataFrame({
    "Category": list(INSIGHTS_DATA[department].keys()),
    "Cost (₹ Lakhs)": list(INSIGHTS_DATA[department].values())
})

st.markdown("### Cost Composition")
st.altair_chart(
    alt.Chart(composition_df)
    .mark_bar(color="#90caf9")
    .encode(
        x=alt.X("Category:N", title="Cost Category", sort="-y"),
        y=alt.Y("Cost (₹ Lakhs):Q", title="Investment Amount (₹ Lakhs)"),
        tooltip=["Category", "Cost (₹ Lakhs)"]
    )
    .properties(height=300),
    use_container_width=True
)

# -------------------------
# Department Comparison
# -------------------------
dept_df = pd.DataFrame({
    "Department": list(DEPARTMENT_TOTALS.keys()),
    "Total Investment (₹ Lakhs)": list(DEPARTMENT_TOTALS.values())
})

st.markdown("### Department-wise Total Investment")
st.altair_chart(
    alt.Chart(dept_df)
    .mark_bar(color="#64b5f6")
    .encode(
        x=alt.X("Department:N", title="Department"),
        y=alt.Y("Total Investment (₹ Lakhs):Q", title="Total Investment (₹ Lakhs)"),
        tooltip=["Department", "Total Investment (₹ Lakhs)"]
    )
    .properties(height=300),
    use_container_width=True
)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Spreadsheet Rendering & Executive Insight Layer"
)
