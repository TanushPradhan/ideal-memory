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
# HTML TABLE RENDERER (FINAL + CENTER ALIGNMENT)
# =========================================================
def render_html_table(df):
    html = """
    <style>
        .excel-container {
            max-height: 75vh;
            overflow: auto;
            border: 1px solid #bfbfbf;
            background-color: #ffffff;
        }

        table {
            border-collapse: collapse;
            width: 100%;
            font-size: 13px;
            font-family: Arial, Helvetica, sans-serif;
            color: #000000;
        }

        td {
            border: 1px solid #c0c0c0;
            padding: 6px 8px;
            vertical-align: top;
            white-space: pre-wrap;
            word-wrap: break-word;
        }

        /* Left align descriptive columns */
        td:nth-child(1),
        td:nth-child(2),
        td:nth-child(9) {
            text-align: left;
        }

        /* Center align numeric / categorical columns */
        td:nth-child(3),
        td:nth-child(4),
        td:nth-child(5),
        td:nth-child(6),
        td:nth-child(7),
        td:nth-child(8) {
            text-align: center;
        }

        tr:nth-child(even) td {
            background-color: #fafafa;
        }
    </style>

    <div class="excel-container">
    <table>
    """

    for _, row in df.iterrows():
        html += "<tr>"
        for cell in row:
            html += f"<td>{cell}</td>"
        html += "</tr>"

    html += "</table></div>"
    return html

# =========================================================
# DISPLAY
# =========================================================
st.subheader(f"📄 Spreadsheet View — {department}")
st.caption("Excel-like, read-only view with wrapped text, gridlines, and centered numeric columns.")

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
