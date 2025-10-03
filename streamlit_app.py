import streamlit as st
import pandas as pd
import plotly.express as px
import json
from pathlib import Path

st.set_page_config(page_title="Vietnam Market Share Report", page_icon="📊", layout="wide")

@st.cache_data(show_spinner=False)
def load_excel(path: str):
    xls = pd.ExcelFile(path)
    sheet = "Data Collection" if "Data Collection" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet)
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
    cols = [c for c in rename_map.keys() if c in df.columns]
    df = df[cols].rename(columns=rename_map)
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    for col in ["Industry", "Parameter/Method", "Instrument Type", "Manufacturer", "Age"]:
        if col in df.columns:
            df[col] = df[col].astype("string").fillna("Unknown")
    return df

def kpi_card(label, value):
    st.metric(label, f"{value:,.0f}")

def group_count(df, by_cols):
    by_cols = [c for c in by_cols if c in df.columns]
    if not by_cols:
        return pd.DataFrame(columns=["Count"])
    return df.groupby(by_cols, dropna=False).size().reset_index(name="Count").sort_values("Count", ascending=False)

st.title("📊 Vietnam Market Share Report")
st.caption("Phân tích tương tác theo **Industry · Parameter/Method · Instrument Type · Manufacturer · Age**")

# Sidebar: file
default_candidates = [Path("data/data.xlsx")]
available_default = next((p for p in default_candidates if p.exists()), None)
uploaded = st.sidebar.file_uploader("Tải file Excel (.xlsx/.xlsm)", type=["xlsx", "xlsm"])

if uploaded is not None:
    df = load_excel(uploaded)
elif available_default is not None:
    df = load_excel(str(available_default))
else:
    st.warning("Hãy upload một file Excel để bắt đầu.")
    st.stop()

# Filters
st.sidebar.header("Bộ lọc")
industries = sorted(df["Industry"].dropna().unique().tolist()) if "Industry" in df.columns else []
instrument_types = sorted(df["Instrument Type"].dropna().unique().tolist()) if "Instrument Type" in df.columns else []
manufacturers = sorted(df["Manufacturer"].dropna().unique().tolist()) if "Manufacturer" in df.columns else []
ages = sorted(df["Age"].dropna().unique().tolist()) if "Age" in df.columns else []

sel_industry = st.sidebar.multiselect("Industry", industries, default=industries[:5] if industries else [])
sel_instr = st.sidebar.multiselect("Instrument Type", instrument_types, default=[])
sel_manu = st.sidebar.multiselect("Manufacturer", manufacturers, default=[])
sel_age = st.sidebar.multiselect("Age", ages, default=[])

df_f = df.copy()
if sel_industry:
    df_f = df_f[df_f["Industry"].isin(sel_industry)]
if sel_instr:
    df_f = df_f[df_f["Instrument Type"].isin(sel_instr)]
if sel_manu:
    df_f = df_f[df_f["Manufacturer"].isin(sel_manu)]
if sel_age:
    df_f = df_f[df_f["Age"].isin(sel_age)]

# KPI
c1, c2, c3, c4 = st.columns(4)
kpi_card("Số bản ghi", len(df_f))
kpi_card("Số ngành", df_f["Industry"].nunique() if "Industry" in df_f.columns else 0)
kpi_card("Loại thiết bị", df_f["Instrument Type"].nunique() if "Instrument Type" in df_f.columns else 0)
kpi_card("Hãng sản xuất", df_f["Manufacturer"].nunique() if "Manufacturer" in df_f.columns else 0)

st.markdown("---")

# Main tabs
tab1, tab2, tab3, tab4 = st.tabs(["Industry", "Instrument Type", "Manufacturer", "Age"])

with tab1:
    st.subheader("Thị phần theo Industry")
    if "Industry" in df_f.columns:
        ind = group_count(df_f, ["Industry"])
        fig = px.pie(ind, names="Industry", values="Count", hole=0.35)
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(ind, use_container_width=True)
    else:
        st.info("Không tìm thấy cột Industry")

with tab2:
    st.subheader("Instrument Type theo Industry")
    need = all(c in df_f.columns for c in ["Industry", "Instrument Type"])
    if need:
        inst = group_count(df_f, ["Industry", "Instrument Type"])
        fig = px.bar(inst, x="Instrument Type", y="Count", color="Industry", barmode="group")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(inst, use_container_width=True)
    else:
        st.info("Thiếu cột Industry/Instrument Type")

with tab3:
    st.subheader("Manufacturer theo Industry")
    need = all(c in df_f.columns for c in ["Industry", "Manufacturer"])
    if need:
        manu = group_count(df_f, ["Industry", "Manufacturer"])
        fig = px.bar(manu, x="Manufacturer", y="Count", color="Industry", barmode="group")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(manu, use_container_width=True)
    else:
        st.info("Thiếu cột Industry/Manufacturer")

with tab4:
    st.subheader("Độ tuổi thiết bị theo Industry")
    need = all(c in df_f.columns for c in ["Industry", "Age"])
    if need:
        ages = group_count(df_f, ["Industry", "Age"])
        fig = px.bar(ages, x="Age", y="Count", color="Industry", barmode="group")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(ages, use_container_width=True)
    else:
        st.info("Thiếu cột Industry/Age")

# --- AI Report Tab (Tiếng Việt) ---
with st.expander("🔐 OpenAI API Key (tạo báo cáo AI)"):
    st.write("App sẽ ưu tiên đọc API key từ **Secrets** hoặc biến môi trường; nếu chạy local có thể nhập tay.")
    import os
    api_key = None
    if "OPENAI_API_KEY" in st.secrets:
        api_key = st.secrets["OPENAI_API_KEY"]
        os.environ["OPENAI_API_KEY"] = api_key
        st.success("🔑 Đã nạp từ Secrets.")
    elif os.getenv("OPENAI_API_KEY"):
        api_key = os.getenv("OPENAI_API_KEY")
        st.success("🔑 Đã nạp từ Environment.")
    else:
        _manual_key = st.text_input("Nhập OPENAI_API_KEY (chỉ khi chạy local)", type="password")
        if _manual_key:
            os.environ["OPENAI_API_KEY"] = _manual_key
            api_key = _manual_key
            st.success("🔑 Đã nhập thủ công.")

ai_tab = st.tabs(["AI Report"])[0]
with ai_tab:
    st.subheader("Sinh báo cáo (Markdown) bằng tiếng Việt")
    st.caption("Ví dụ: 'Viết executive summary cho Industrial & Water Treatment, tập trung Controller/Benchtop, nhấn mạnh nhóm 3–5 năm'")
    user_req = st.text_area(
        "Nhập yêu cầu báo cáo (Tiếng Việt)",
        height=120,
        placeholder="Ví dụ: Tạo báo cáo 1 trang tổng quan top 5 ngành theo số lượng thiết bị..."
    )
    model_choice = st.selectbox("Model", ["gpt-4o-mini", "gpt-4o"], index=0)
    run = st.button("Tạo báo cáo")
    if run and user_req.strip():
        try:
            from ai_report import generate_report
            md = generate_report(df_f if len(df_f) else df, user_req, model=model_choice)
            st.markdown(md)
            st.download_button("Tải Markdown", md.encode("utf-8"), file_name="market_report.md", mime="text/markdown")
        except Exception as e:
            st.error(f"Lỗi tạo báo cáo: {e}")

# --- AI Charts (Tạo biểu đồ từ tiếng Việt) ---
charts_tab = st.tabs(["AI Charts"])[0]
with charts_tab:
    st.subheader("Vẽ biểu đồ theo yêu cầu (Tiếng Việt)")
    st.caption("Ví dụ: 'Vẽ biểu đồ tròn thị phần theo Industry, top 5 ngành' hoặc 'Vẽ biểu đồ cột theo Manufacturer cho ngành Water Treatment'")

    with st.expander("🔐 OpenAI API Key (tạo biểu đồ AI)"):
        import os
        api_key2 = None
        if "OPENAI_API_KEY" in st.secrets:
            api_key2 = st.secrets["OPENAI_API_KEY"]
            os.environ["OPENAI_API_KEY"] = api_key2
            st.success("Đã nạp API key từ Secrets.")
        elif os.getenv("OPENAI_API_KEY"):
            api_key2 = os.getenv("OPENAI_API_KEY")
            st.success("Đã nạp API key từ Environment.")
        else:
            _manual_key = st.text_input("Nhập OPENAI_API_KEY (local)", type="password", key="chart_key")
            if _manual_key:
                os.environ["OPENAI_API_KEY"] = _manual_key
                api_key2 = _manual_key
                st.success("API key đã được nhập.")

    user_chart_req = st.text_area("Mô tả biểu đồ mong muốn", height=100, placeholder="Ví dụ: Vẽ biểu đồ tròn thị phần theo Industry, top 5 ngành")
    model_chart = st.selectbox("Model cho AI Charts", ["gpt-4o-mini", "gpt-4o"], index=0)
    run_chart = st.button("Tạo biểu đồ bằng AI")
    if run_chart and user_chart_req.strip():
        try:
            from ai_charts import infer_spec_via_openai, render_chart
            spec = infer_spec_via_openai(user_chart_req, model=model_chart)
            st.code(json.dumps(spec, ensure_ascii=False, indent=2), language="json")
            fig, data_used = render_chart(df_f if len(df_f) else df, spec)
            st.plotly_chart(fig, use_container_width=True)
            st.dataframe(data_used, use_container_width=True)
        except Exception as e:
            st.error(f"Không thể tạo biểu đồ: {e}")

st.markdown("---")
st.caption("Mẹo: Dùng sidebar để lọc dữ liệu; Upload file Excel mới để cập nhật phân tích.")