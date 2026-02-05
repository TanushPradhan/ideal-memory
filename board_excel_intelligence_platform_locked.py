import streamlit as st
import pandas as pd
import os
import matplotlib.pyplot as plt
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
    "Board-ready dashboard for Academic Expansion Plans "
    "(Gemology, Manufacturing & CAD)"
)

# =========================================================
# FILE CONFIG
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
    return pd.read_excel(path, sheet_name=sheet).fillna("")

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

# Remove numeric-only headers like 60
df = df.loc[:, ~df.columns.map(lambda x: str(x).strip().isdigit())]

df.rename(columns=COLUMN_NAME_OVERRIDES.get(department, {}), inplace=True)
df_display = df.fillna("")

# =========================================================
# AGGRID SETUP (unchanged)
# =========================================================
gb = GridOptionsBuilder.from_dataframe(df_display)
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
        df_display,
        gridOptions=gb.build(),
        update_mode=GridUpdateMode.NO_UPDATE,
        allow_unsafe_jscode=True,
        theme="alpine",
        height=720
    )

# =========================================================
# EXECUTIVE VIEW WITH INSIGHTS
# =========================================================
else:
    st.subheader(f"🧑‍💼 Executive View — {department}")
    st.caption("Static insights derived directly from the provided spreadsheet.")

    # -----------------------------------------------------
    # Identify numeric columns
    # -----------------------------------------------------
    numeric_df = df.apply(pd.to_numeric, errors="coerce")
    numeric_cols = numeric_df.columns[numeric_df.notna().any()].tolist()

    # -----------------------------------------------------
    # KPI CARDS
    # -----------------------------------------------------
    total_rows = len(df)
    total_numeric_sum = numeric_df.sum().sum()

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Line Items", total_rows)
    c2.metric("Numeric Columns", len(numeric_cols))
    c3.metric("Total Financial Outlay (₹ Lakhs)", f"{total_numeric_sum:,.2f}")

    st.markdown("---")

    # -----------------------------------------------------
    # COST COLUMN DETECTION (SAFE HEURISTIC)
    # -----------------------------------------------------
    cost_cols = [c for c in numeric_cols if "cost" in c.lower() or "total" in c.lower()]

    if cost_cols:
        cost_col = cost_cols[-1]

        cost_series = numeric_df[cost_col].dropna()
        top_costs = cost_series.sort_values(ascending=False).head(5)

        st.subheader("🔍 Key Insights")

        st.markdown(
            f"""
            • The **total projected expenditure** for the {department} department is
              **₹ {cost_series.sum():,.2f} Lakhs**.  
            • The **top 5 cost items** contribute approximately
              **{(top_costs.sum() / cost_series.sum()) * 100:.1f}%** of the total spend.  
            • Cost concentration suggests a **capital-intensive structure** driven by
              high-value equipment.
            """
        )

        # -------------------------------------------------
        # BAR CHART
        # -------------------------------------------------
        fig, ax = plt.subplots()
        top_costs.plot(kind="bar", ax=ax)
        ax.set_title("Top Cost Contributors")
        ax.set_ylabel("Cost (₹ Lakhs)")
        ax.set_xlabel("Line Item")
        plt.tight_layout()

        st.pyplot(fig)

        # -------------------------------------------------
        # PIE CHART
        # -------------------------------------------------
        fig2, ax2 = plt.subplots()
        top_costs.plot(kind="pie", ax=ax2, autopct="%1.1f%%")
        ax2.set_ylabel("")
        ax2.set_title("Share of Total Spend (Top 5 Items)")
        plt.tight_layout()

        st.pyplot(fig2)

    st.markdown("---")
    st.subheader("📄 Executive Snapshot (Reference)")
    AgGrid(
        df_display,
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
