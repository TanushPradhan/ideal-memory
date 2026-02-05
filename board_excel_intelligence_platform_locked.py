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
# SAFE LOADER
# =========================================================
def load_excel(path, sheet):
    if not os.path.exists(path):
        st.error(f"File not found: {path}")
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
config = DEPARTMENT_CONFIG[department]
df = load_excel(config["file"], config["sheet"])
df_display = df.fillna("")

# =========================================================
# EXECUTIVE INSIGHTS (STATIC & SAFE)
# =========================================================
def render_insights(dept):
    st.markdown("## 🧠 Executive Insights")
    st.markdown("---")

    if dept == "Gemology":
        st.markdown("""
### 📌 Program Structure
- **Planned intake:** ~60–65 students  
- **Spare buffer:** 2–3 additional instruments for critical items  
- **Faculty ratio:** 1 faculty per 25 students (≈ **2.4 faculty**)

### 💎 Cost Characteristics
- High-value instruments (Microscopes, Optical tools) are **shared**
- Per-student accessories are **low cost, high volume**
- Clear separation between **per-student** and **shared resources**

### 🏛 Academic & Compliance Strength
- Room sealing and lighting conditions explicitly specified
- Equipment selection aligns with **international gemology standards**

**Board View:**  
> Capital-efficient, compliance-driven expansion with long asset life.
""")

    elif dept == "Manufacturing":
        st.markdown("""
### 📌 Financial Scale
- **Total estimated cost:** ₹ **166.03 Lakhs**
- Most capital-intensive department

### 🏗 Major Cost Drivers
- Casting machines, rolling machines, furnaces
- Dust collectors, polishing systems, safety infrastructure

### ⚙ Utilisation Logic
- Majority of machinery **shared across batches**
- Capacity calculations clearly documented in remarks

### ⚠ Risk Perspective
- High capex dependency
- Expansion beyond current capacity will require reinvestment

**Board View:**  
> High capital requirement, but technically unavoidable and operationally justified.
""")

    elif dept == "CAD":
        st.markdown("""
### 📌 Cost Nature
- Lower capital expenditure compared to Manufacturing
- Primary costs driven by **systems and software licenses**

### 🔁 Cost Behaviour
- Hardware: replaceable
- Software: recurring (license renewals, upgrades)

### 📈 Scalability
- Easiest department to scale
- Minimal infrastructure constraints

**Board View:**  
> Most flexible and scalable vertical with controlled financial risk.
""")

# =========================================================
# RENDER
# =========================================================
if view_mode == "Executive View":
    render_insights(department)
    st.markdown("---")

st.subheader(f"📄 Spreadsheet View — {department}")
st.caption("Excel-like, read-only view. All figures visible. No assumptions.")

st.dataframe(
    df_display,
    use_container_width=True,
    height=650
)

# =========================================================
# FOOTER
# =========================================================
st.markdown("---")
st.caption("© Board Excel Intelligence Platform — Academic Expansion Dashboard")
