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
    "Locked, board-ready spreadsheet view for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# DEPARTMENT CONFIG (FILES ARE SOURCE OF TRUTH)
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
# SAFE EXCEL LOADER (NO DATA LOSS)
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"Required file not found:\n{path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet, header=None)

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

# Replace NaN with empty string (DO NOT DROP ROWS)
df = df.fillna("")

# =========================================================
# RENDER AS TRUE HTML TABLE (GRIDLINES + WRAP)
# =========================================================
def render_table(df):
    html = """
    <style>
        .excel-table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
        }
        .excel-table th,
        .excel-table td {
            border: 1px solid #d0d0d0;
            padding: 6px 8px;
            vertical-align: top;
            text-align: left;
            white-space: pre-wrap;
            word-wrap: break-word;
        }
        .excel-table th {
            background-color: #f5f5f5;
            font-weight: 600;
        }
        .excel-container {
            max-height: 75vh;
            overflow: auto;
        }
    </style>
    <div class="excel-container">
    <table class="excel-table">
    """

    # Render rows exactly as-is
    for r_idx, row in df.iterrows():
        html += "<tr>"
        for cell in row:
            tag = "th" if r_idx == 0 else "td"
            html += f"<{tag}>{cell}</{tag}>"
        html += "</tr>"

    html += "</table></div>"
    return html

# =========================================================
# DISPLAY
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption("Excel-like, read-only view with wrapped text and full gridlines.")

st.markdown(render_table(df), unsafe_allow_html=True)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Spreadsheet Rendering Layer"
)
