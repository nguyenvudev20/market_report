# Vietnam Market Share Report (Streamlit)

An interactive Streamlit app for analyzing market data by **Industry · Parameter/Method · Instrument Type · Manufacturer · Age**.

## 🗂️ Project structure
```
streamlit_market_report/
├─ streamlit_app.py
├─ requirements.txt
├─ README.md
└─ data/
   └─ Market_Analysis_Report.xlsx   # Replace or add your original .xlsm if preferred
```

## 🚀 Run locally
```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

Then open the URL shown in your terminal.

## ☁️ Deploy on Streamlit Community Cloud
1. Create a new **GitHub** repo (public).
2. Upload all files from `streamlit_market_report/` to the repo.
3. Go to https://share.streamlit.io/ → New app → Connect your repo.
4. Set **Main file path** to `streamlit_app.py` and deploy.

> If your dataset is large or private, consider storing it in a private repo and/or using Streamlit **secrets** to access from cloud storage (e.g., S3, GDrive API).

## 📦 Use your original Excel
- Put your Excel file into `data/` and name it anything (e.g., `Vietnam_6789_2025_Market Share Report.xlsm`).
- In the app, use the **Upload** widget in the sidebar, or replace the default `Market_Analysis_Report.xlsx` inside `data/`.

## ✨ Features
- Filters by Industry, Instrument Type, Manufacturer, and Age
- KPI counters and interactive charts (Pie and Bar)
- Works with both `.xlsx` and `.xlsm` (macros are ignored by pandas)