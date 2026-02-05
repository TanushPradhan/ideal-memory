import streamlit as st
import pandas as pd
import os
import matplotlib.pyplot as plt
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode

# =========================================================
# APP CONFIG
# =========================================================
st.set_page_config(
    page_title="Board Excel Intelligence Platform",
    layout="wide"
)

st.title("📊 Board Excel Intelligence Platform")
st.caption(
    "Board-ready dashboard for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# FILE CONFIG
# =========================================================
DEPARTMENT_CONFIG = {
    "Gemology": {
        "file": "IIGJ Mumbai - Academic Expansion Plan - Gemology Formatted.xlsx",
        "sheet": "Department of Gemmology_Format",
        "cost_col": "Total in Lakhs",
        "qty_col": "Quantity"
    },
    "Manufacturing": {
        "file": "IIGJ - Academic Expansion Plan - Manufacturing Formatted.xlsx",
        "sheet": "Formated_Budget",
        "cost_col": "Cost-to-Company (in Lakhs)",
        "qty_col": "Quantity"
    },
    "CAD": {
        "file": "IIGJ Mumbai - Academic expansion Plan_CAD Formatted.xlsx",
        "sheet": "Formatted_Budget",
        "cost_col": "Total Cost",
        "qty_col": "Quantity"
    }
}

# =========================================================
# COLUMN HEADERS
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
        "Unnamed: 6": "According to Capacity",
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
# SAFE LOADER
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"Missing file: {path}")
        st.stop()
    return pd.read_excel(path, sheet_name=sheet)

# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.header("📁 Navigation")

department = st.sidebar.selectbox(
    "Select Department",
    list(DEPARTMENT_CONFIG.keys())
)

view_mode = st.sidebar.radio(
    "View Mode",
    ["Interactive Spreadsheet", "Executive View"]
)

# =========================================================
# LOAD DATA
# =========================================================
cfg = DEPARTMENT_CONFIG[department]
df = load_excel(cfg["file"], cfg["sheet"])

# Remove stray numeric headers like 60
df = df.loc[:, ~df.columns.map(lambda x: str(x).strip().isdigit())]

df.rename(columns=COLUMN_NAME_OVERRIDES.get(department, {}), inplace=True)
df = df.fillna("")

# =========================================================
# GRID (UNCHANGED)
# =========================================================
gb = GridOptionsBuilder.from_dataframe(df)
gb.configure_default_column(
    wrapText=True,
    autoHeight=True,
    resizable=True,
    filter=True,
    sortable=False,
    cellStyle={
        "borderRight": "1px solid #D0D0D0",
        "borderBottom": "1px solid #D0D0D0",
        "whiteSpace": "normal",
        "lineHeight": "1.4"
    }
)
gb.configure_grid_options(enableRangeSelection=True, rowHeight=44)

# =========================================================
# INTERACTIVE VIEW
# =========================================================
if view_mode == "Interactive Spreadsheet":
    st.subheader(f"📄 Interactive Spreadsheet — {department}")

    AgGrid(
        df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,
        allow_unsafe_jscode=True,
        theme="alpine",
        height=720
    )

# =========================================================
# EXECUTIVE VIEW (FIXED FIGURES)
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {department}")
    st.caption("All figures are computed directly from the approved spreadsheet.")

    # ---- Explicit numeric conversion ----
    cost_col = cfg["cost_col"]
    qty_col = cfg["qty_col"]

    df[cost_col] = pd.to_numeric(df[cost_col], errors="coerce")
    if qty_col in df.columns:
        df[qty_col] = pd.to_numeric(df[qty_col], errors="coerce")

    total_budget = df[cost_col].sum()
    total_qty = df[qty_col].sum() if qty_col in df.columns else None
    total_items = df[cost_col].notna().sum()

    # ---- KPI CARDS ----
    c1, c2, c3 = st.columns(3)
    c1.metric("Total Line Items", total_items)
    c2.metric("Total Quantity", f"{int(total_qty)}" if total_qty is not None else "—")
    c3.metric("Grand Total Budget (₹ Lakhs)", f"{total_budget:,.2f}")

    st.markdown("---")

    # ---- TOP COST DRIVERS ----
    top_costs = df[[cost_col]].dropna().sort_values(cost_col, ascending=False).head(5)

    st.subheader("🔍 Top Cost Drivers")
    st.dataframe(top_costs, use_container_width=True)

    # ---- BAR CHART ----
    fig, ax = plt.subplots()
    top_costs[cost_col].plot(kind="bar", ax=ax)
    ax.set_title("Top 5 Cost Contributors")
    ax.set_ylabel("Cost (₹ Lakhs)")
    ax.set_xlabel("Line Item Index")
    plt.tight_layout()
    st.pyplot(fig)

    # ---- PIE CHART ----
    fig2, ax2 = plt.subplots()
    top_costs[cost_col].plot(kind="pie", ax=ax2, autopct="%1.1f%%")
    ax2.set_ylabel("")
    ax2.set_title("Share of Total Spend (Top 5)")
    plt.tight_layout()
    st.pyplot(fig2)

    st.markdown("---")
    st.subheader("📄 Executive Reference Sheet")

    AgGrid(
        df,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,
        allow_unsafe_jscode=True,
        theme="alpine",
        height=500
    )

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption("© Board Excel Intelligence Platform — Executive Intelligence Layer")
