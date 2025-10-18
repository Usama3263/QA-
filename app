import streamlit as st
import pandas as pd
import numpy as np
import os
import io
import glob

# ------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------

def clean_data(df):
    df.columns = [str(col).replace(" [%]", "").strip().replace("\xa0", "") for col in df.columns]
    id_cols = ["Name", "Gerät", "Nummer", "Typ", "Mandant"]

    df = df[df["Typ"].isin(["FV", "LA", "RA"])]
    score_cols = [c for c in df.columns if c not in id_cols]

    df[score_cols] = df[score_cols].apply(pd.to_numeric, errors='coerce')
    df[score_cols] = df[score_cols].replace([-102, 0], np.nan)

    df_long = df.melt(id_vars=["Gerät", "Typ"], value_vars=score_cols,
                      var_name="Day", value_name="Score")
    df_long = df_long[df_long["Score"].notna()]
    return df_long

def analyze_month(df_long):
    summary = df_long.groupby(["Gerät", "Typ"])["Score"].agg(
        Average_Score="mean",
        P90=lambda x: np.percentile(x, 90)
    ).reset_index()
    return summary

def build_annual_dashboard(output_folder):
    all_summaries = [f for f in os.listdir(output_folder) if f.endswith("_summary.xlsx")]
    combined = pd.DataFrame()
    for file in sorted(all_summaries):
        path = os.path.join(output_folder, file)
        month_part = os.path.splitext(file)[0].replace("_summary", "")
        try:
            year, month = month_part.split(".")
            month_label = f"{month}.{year}"
        except ValueError:
            month_label = month_part
        df = pd.read_excel(path)
        if "Average_Score" not in df.columns:
            continue
        avg_by_device = df.groupby("Gerät")["Average_Score"].mean().reset_index()
        avg_by_device.rename(columns={"Average_Score": month_label}, inplace=True)
        if combined.empty:
            combined = avg_by_device
        else:
            combined = pd.merge(combined, avg_by_device, on="Gerät", how="outer")
    return combined

def style_scores(df, score_cols):
    def cell_color(val):
        if pd.isna(val):
            return ""
        elif val >= 90:
            return "background-color: #b7e1cd"
        elif val >= 80:
            return "background-color: #fff2cc"
        else:
            return ""
    styled = df.style.applymap(cell_color, subset=score_cols)
    styled = styled.set_properties(**{'font-weight': 'bold', 'text-transform': 'uppercase'}, subset=df.columns)
    return styled

# ------------------------------------------------------------
# Streamlit App
# ------------------------------------------------------------

st.set_page_config(page_title="QA Analysis Dashboard", layout="wide")
st.title("📊 QUALITY ASSURANCE (QA) DATA ANALYSIS DASHBOARD")

output_folder = "QA_Analysis"
os.makedirs(output_folder, exist_ok=True)

tab1, tab2 = st.tabs(["Monthly Analysis", "Annual Dashboard"])

with tab1:
    st.header("MONTHLY QA ANALYSIS")
    uploaded_files = st.file_uploader(
        "Upload one or more monthly QA Excel files",
        type=["xlsx"],
        accept_multiple_files=True
    )

    if uploaded_files:
        for f in glob.glob(os.path.join(output_folder, "*_summary.xlsx")):
            os.remove(f)
        for uploaded_file in uploaded_files:
            st.write(f"📂 Processing file: {uploaded_file.name}")
            df = pd.read_excel(uploaded_file)
            df_long = clean_data(df)
            summary = analyze_month(df_long)
            st.subheader(f"Results for {uploaded_file.name}")
            score_cols = ["Average_Score", "P90"]
            st.dataframe(style_scores(summary, score_cols))
            base_name = os.path.splitext(uploaded_file.name)[0]
            summary_path = os.path.join(output_folder, f"{base_name}_summary.xlsx")
            summary.to_excel(summary_path, index=False)
            st.info(f"💾 Saved monthly summary: {summary_path}")
        st.success("✅ All uploaded files processed successfully!")

with tab2:
    st.header("ANNUAL QA DASHBOARD")
    dashboard = build_annual_dashboard(output_folder)
    if dashboard.empty:
        st.warning("⚠️ No summary data yet. Upload monthly files in Tab 1 first.")
    else:
        score_cols = [c for c in dashboard.columns if c != "Gerät"]
        st.subheader("Annual QA Table (Sortable)")
        st.dataframe(style_scores(dashboard, score_cols), use_container_width=True)
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            dashboard.to_excel(writer, index=False, sheet_name="Annual_Dashboard")
        buffer.seek(0)
        st.download_button(
            label="⬇️ Download Annual Dashboard (XLSX)",
            data=buffer,
            file_name="Annual_Dashboard.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
