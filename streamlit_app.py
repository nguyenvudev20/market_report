import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

st.set_page_config(page_title="Vietnam Market Share Report", page_icon="📊", layout="wide")

@st.cache_data(show_spinner=False)
def load_excel(path: str):
    # Try multiple sheet names gracefully
    xls = pd.ExcelFile(path)
    sheet = "Data Collection" if "Data Collection" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
    # Normalize column names we care about
    rename_map = {
        "Age of Product": "Age",
        "Instrument Type": "Instrument Type",
        "Parameter/Method": "Parameter/Method",
        "Manufacturer": "Manufacturer",
        "Industry": "Industry",
        "Date": "Date",
        "Hanna Office": "Hanna Office",
        "Hanna Rep": "Hanna Rep",
        "Customer Name": "Customer Name",
        "Model #": "Model"
    }
    # Keep only columns present
    cols = [c for c in rename_map.keys() if c in df.columns]
    df = df[cols].rename(columns=rename_map)
    # Ensure types
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for col in ["Industry", "Parameter/Method", "Instrument Type", "Manufacturer", "Age"]:
        if col in df.columns:
            df[col] = df[col].astype("string").fillna("Unknown")
    return df

def kpi_card(label, value):
    st.metric(label, f"{value:,.0f}")

def group_count(df, by_cols):
    return df.groupby(by_cols, dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)

st.title("📊 Vietnam Market Share Report")
st.caption("Interactive analysis by **Industry · Parameter/Method · Instrument Type · Manufacturer · Age**")

# Sidebar: file selector & filters
default_path = Path("data/Market_Analysis_Report.xlsx")
uploaded = st.sidebar.file_uploader("Upload Excel file (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

if uploaded is not None:
    df = load_excel(uploaded)
elif default_path.exists():
    df = load_excel(str(default_path))
else:
    st.warning("Please upload an Excel file to begin.")
    st.stop()

st.sidebar.header("Filters")
industries = sorted(df["Industry"].dropna().unique().tolist()) if "Industry" in df.columns else []
instrument_types = sorted(df["Instrument Type"].dropna().unique().tolist()) if "Instrument Type" in df.columns else []
manufacturers = sorted(df["Manufacturer"].dropna().unique().tolist()) if "Manufacturer" in df.columns else []
ages = sorted(df["Age"].dropna().unique().tolist()) if "Age" in df.columns else []

sel_industry = st.sidebar.multiselect("Industry", industries, default=industries[:5] if industries else [])
sel_instr = st.sidebar.multiselect("Instrument Type", instrument_types, default=[])
sel_manu = st.sidebar.multiselect("Manufacturer", manufacturers, default=[])
sel_age = st.sidebar.multiselect("Age of Product", ages, default=[])

df_f = df.copy()
if sel_industry:
    df_f = df_f[df_f["Industry"].isin(sel_industry)]
if sel_instr:
    df_f = df_f[df_f["Instrument Type"].isin(sel_instr)]
if sel_manu:
    df_f = df_f[df_f["Manufacturer"].isin(sel_manu)]
if sel_age:
    df_f = df_f[df_f["Age"].isin(sel_age)]

# KPI row
c1, c2, c3, c4 = st.columns(4)
kpi_card("Records", len(df_f))
if "Industry" in df_f.columns:
    kpi_card("Industries", df_f["Industry"].nunique())
else:
    c2.write("")
if "Instrument Type" in df_f.columns:
    kpi_card("Instrument Types", df_f["Instrument Type"].nunique())
else:
    c3.write("")
if "Manufacturer" in df_f.columns:
    kpi_card("Manufacturers", df_f["Manufacturer"].nunique())
else:
    c4.write("")

st.markdown("---")

# Charts
tab1, tab2, tab3, tab4 = st.tabs(["Industry", "Instrument Type", "Manufacturer", "Age"])

with tab1:
    st.subheader("Industry Market Share")
    if "Industry" in df_f.columns:
        ind = group_count(df_f, ["Industry"])
        fig = px.pie(ind, names="Industry", values="Count", hole=0.35)
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(ind, use_container_width=True)
    else:
        st.info("Industry column not found in the dataset.")

with tab2:
    st.subheader("Instrument Type by Industry")
    needed = all(c in df_f.columns for c in ["Industry", "Instrument Type"])
    if needed:
        inst = group_count(df_f, ["Industry", "Instrument Type"])
        fig = px.bar(inst, x="Instrument Type", y="Count", color="Industry", barmode="group")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(inst.head(200), use_container_width=True)
    else:
        st.info("Columns missing: Industry and/or Instrument Type.")

with tab3:
    st.subheader("Manufacturer by Industry")
    needed = all(c in df_f.columns for c in ["Industry", "Manufacturer"])
    if needed:
        manu = group_count(df_f, ["Industry", "Manufacturer"])
        fig = px.bar(manu, x="Manufacturer", y="Count", color="Industry", barmode="group")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(manu.head(200), use_container_width=True)
    else:
        st.info("Columns missing: Industry and/or Manufacturer.")

with tab4:
    st.subheader("Age Distribution by Industry")
    needed = all(c in df_f.columns for c in ["Industry", "Age"])
    if needed:
        ages = group_count(df_f, ["Industry", "Age"])
        fig = px.bar(ages, x="Age", y="Count", color="Industry", barmode="group")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(ages.head(200), use_container_width=True)
    else:
        st.info("Columns missing: Industry and/or Age.")

st.markdown("---")
st.caption("Tip: Use the sidebar to filter by Industry, Instrument Type, Manufacturer, and Age. Upload a new Excel file to refresh the analysis.")