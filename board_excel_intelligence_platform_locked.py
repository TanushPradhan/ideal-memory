import streamlit as st
import pandas as pd
import os
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode

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
# DEPARTMENT CONFIG
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
# PROFESSIONAL COLUMN NAME OVERRIDES
# =========================================================
COLUMN_NAME_OVERRIDES = {
    "Gemology": {
        "Unnamed: 1": "Instrument Name",
        "Unnamed: 2": "Type",
        "Unnamed: 3": "Specification",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Unit Price (in Lakhs)",
        "Unnamed: 6": "Total in Lakhs",
        "Unnamed: 7": "Remarks"
    },
    "Manufacturing": {
        "Unnamed: 1": "Particulars",
        "Unnamed: 2": "Department",
        "Unnamed: 3": "Details",
        "Unnamed: 4": "Quantity",
        "Unnamed: 5": "Cost per Item (in Lakhs)",
        "Unnamed: 6": "According to Capacity (60 Students)",
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
        st.error(f"Required file not found: {path}")
        st.stop()
    try:
        return pd.read_excel(path, sheet_name=sheet)
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

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
enable_highlight = st.sidebar.checkbox("Enable highlighting", value=False)
highlight_color = st.sidebar.color_picker("Highlight color", "#FFF3A0")

# =========================================================
# SESSION STATE FOR HIGHLIGHTS
# =========================================================
if "highlighted_ranges" not in st.session_state:
    st.session_state.highlighted_ranges = []

# =========================================================
# LOAD DATA
# =========================================================
config = DEPARTMENT_CONFIG[selected_department]
df = load_excel_safely(config["file"], config["sheet"])

# Remove accidental numeric headers like "60"
df.columns = [str(c) if not isinstance(c, (int, float)) else "" for c in df.columns]

# Apply professional column names
df.rename(columns=COLUMN_NAME_OVERRIDES.get(selected_department, {}), inplace=True)

df_display = df.fillna("")

# =========================================================
# AG-GRID CELL STYLE
# =========================================================
cell_style_jscode = JsCode("""
function(params) {
    const ranges = window.highlightedRanges || [];
    for (let i = 0; i < ranges.length; i++) {
        const r = ranges[i];
        if (
            params.rowIndex >= r.startRow &&
            params.rowIndex <= r.endRow &&
            r.columns.includes(params.colDef.field)
        ) {
            return {
                backgroundColor: '%s',
                fontWeight: 'bold'
            };
        }
    }
    return null;
}
""" % highlight_color)

# =========================================================
# GRID BUILDER
# =========================================================
gb = GridOptionsBuilder.from_dataframe(df_display)

gb.configure_default_column(
    wrapText=True,
    autoHeight=True,
    resizable=True,
    sortable=False,
    filter=True,
    cellStyle=cell_style_jscode
)

gb.configure_grid_options(
    enableRangeSelection=True,
    suppressMultiRangeSelection=False,
    rowHeight=42,
    getContextMenuItems=JsCode("""
    function(params) {
        const items = params.defaultItems || [];
        items.push({
            name: 'Highlight Selection',
            action: function() {
                const ranges = params.api.getCellRanges();
                if (!ranges) return;

                window.highlightedRanges = window.highlightedRanges || [];
                ranges.forEach(r => {
                    window.highlightedRanges.push({
                        startRow: Math.min(r.startRow.rowIndex, r.endRow.rowIndex),
                        endRow: Math.max(r.startRow.rowIndex, r.endRow.rowIndex),
                        columns: r.columns.map(c => c.colId)
                    });
                });
            }
        });

        items.push({
            name: 'Clear Highlights',
            action: function() {
                window.highlightedRanges = [];
            }
        });

        return items;
    }
    """)
)

# =========================================================
# RENDER
# =========================================================
st.subheader(
    f"📄 Spreadsheet View — {selected_department}"
    if view_mode == "Interactive Spreadsheet"
    else f"🧑‍💼 Executive View — {selected_department}"
)

AgGrid(
    df_display,
    gridOptions=gb.build(),
    update_mode=GridUpdateMode.NO_UPDATE,
    allow_unsafe_jscode=True,
    theme="alpine",
    height=700,
    fit_columns_on_grid_load=False
)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption(
    "© Board Excel Intelligence Platform — Locked Academic Expansion Dashboard"
)
