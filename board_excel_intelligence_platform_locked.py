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
    "instrument name",
    "type",
    "specification",
    "quantity",
    "unit price",
    "total in lakhs",
    "remarks"
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
# SAFE EXCEL LOADER (DO NOT TOUCH)
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
# HTML TABLE RENDERER (FINAL – DO NOT CHANGE)
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
# TABLE DISPLAY (LOCKED)
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
# EXECUTIVE INSIGHTS (STATIC, SAFE, SEPARATE)
# =========================================================
st.markdown("---")
st.subheader("📊 Executive Insights (Summary)")
st.caption("Strategic overview based on approved expansion plans.")

INSIGHTS_DATA = {
    "Gemology": {
        "Per Student Equipment": 48.2,
        "Shared Equipment": 72.4,
        "Consumables": 18.6,
        "Infrastructure": 12.8
    },
    "Manufacturing": {
        "Heavy Machinery": 92.5,
        "Shared Equipment": 48.3,
        "Per Student Equipment": 21.2,
        "Consumables": 4.0
    },
    "CAD": {
        "Systems & Licenses": 38.0,
        "Per Student Hardware": 22.0,
        "Shared Infrastructure": 10.0
    }
}

DEPARTMENT_TOTALS = {
    "Gemology": 152.0,
    "Manufacturing": 166.0,
    "CAD": 70.0
}

# ---------------------------------------------------------
# GRAPH 1 — COST COMPOSITION
# ---------------------------------------------------------
st.markdown("### Cost Composition")

comp_df = (
    pd.DataFrame.from_dict(
        INSIGHTS_DATA[department],
        orient="index",
        columns=["Cost (₹ Lakhs)"]
    )
    .reset_index()
    .rename(columns={"index": "Category"})
)

st.bar_chart(comp_df.set_index("Category"), height=280)

# ---------------------------------------------------------
# GRAPH 2 — SHARED VS PER STUDENT
# ---------------------------------------------------------
if department in ["Gemology", "Manufacturing"]:
    st.markdown("### Shared vs Per-Student Cost Split")

    shared = sum(v for k, v in INSIGHTS_DATA[department].items() if "shared" in k.lower())
    per_student = sum(v for k, v in INSIGHTS_DATA[department].items() if "student" in k.lower())

    split_df = pd.DataFrame({
        "Category": ["Shared Costs", "Per-Student Costs"],
        "Cost (₹ Lakhs)": [shared, per_student]
    })

    st.bar_chart(split_df.set_index("Category"), height=240)

# ---------------------------------------------------------
# GRAPH 3 — DEPARTMENT COMPARISON
# ---------------------------------------------------------
st.markdown("### Department-wise Total Investment")

dept_df = pd.DataFrame({
    "Department": list(DEPARTMENT_TOTALS.keys()),
    "Total Cost (₹ Lakhs)": list(DEPARTMENT_TOTALS.values())
})

st.bar_chart(dept_df.set_index("Department"), height=260)

# ---------------------------------------------------------
# EXECUTIVE NOTES (STATIC TEXT)
# ---------------------------------------------------------
st.markdown("### Key Observations")

if department == "Gemology":
    st.markdown("""
- Balanced investment between **shared lab infrastructure** and **per-student instruments**.
- Emphasis on grading accuracy, microscopy, and controlled environments.
- Costs scale moderately with intake due to shared usage models.
""")

elif department == "Manufacturing":
    st.markdown("""
- Capital-heavy department driven by **core machinery**.
- High shared utilization reduces long-term marginal cost per student.
- Infrastructure investment supports multi-batch scalability.
""")

elif department == "CAD":
    st.markdown("""
- Technology-centric expenditure dominated by licenses and systems.
- Predictable per-student scaling.
- Lowest infrastructure risk among departments.
""")

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption("© Board Excel Intelligence Platform — Executive Intelligence Layer")
