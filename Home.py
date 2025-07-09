import streamlit as st
import pandas as pd
from datetime import datetime
import re
from io import BytesIO

st.set_page_config(page_title="PunchPro with Save Button", layout="wide")
st.title("📂 PunchPro")

st.write(
    "Upload one or more Excel/CSV files below. "
    "Each will open in its own tab with cleaned data. "
    "Click **Save Cleaned File** to download in a consistent format."
)

uploaded_files = st.file_uploader(
    "Select one or more files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True
)

if uploaded_files:
    tabs = st.tabs([file.name for file in uploaded_files])

    for idx, file in enumerate(uploaded_files):
        with tabs[idx]:
            st.markdown(f"### 📄 `{file.name}`")

            try:
                # Read raw file
                if file.name.endswith(".csv"):
                    df_raw = pd.read_csv(file, header=None)
                else:
                    df_raw = pd.read_excel(file, header=None)

                # Clean the file
                df_clean = df_raw.iloc[2:].reset_index(drop=True)
                df_clean.columns = df_clean.iloc[0]
                df_clean = df_clean[1:].reset_index(drop=True)

                # Drop irrelevant columns
                if 'S.No' in df_clean.columns:
                    df_clean = df_clean.drop(columns=['S.No'])

                df_clean = df_clean.loc[:, ~df_clean.columns.isna()]

                # 🔷 Extract Time In / Time Out from Punch Records
                def extract_time_in_out(record):
                    if pd.isna(record):
                        return pd.Series(["", ""])
                    punches = re.findall(r'\d{1,2}:\d{2}', str(record))
                    if not punches:
                        return pd.Series(["", ""])
                    time_in = punches[0]
                    time_out = punches[-1]
                    return pd.Series([time_in, time_out])

                df_clean[['Time In', 'Time Out']] = df_clean['Punch Records'].apply(extract_time_in_out)
                df_clean = df_clean.drop(columns=['Punch Records'])

                # 🔷 Calculate Stay Duration
                def calc_stay_duration(row):
                    try:
                        in_time = pd.to_datetime(row['Time In'], format='%H:%M')
                        out_time = pd.to_datetime(row['Time Out'], format='%H:%M')
                        duration = out_time - in_time
                        total_minutes = duration.total_seconds() / 60
                        hours = int(total_minutes // 60)
                        minutes = int(total_minutes % 60)
                        return f"{hours:02}:{minutes:02}"
                    except:
                        return ""

                df_clean['Stay Duration'] = df_clean.apply(calc_stay_duration, axis=1)

                # Show cleaned data
                st.dataframe(df_clean, use_container_width=True)

                # 🔷 Save button (inside each tab)
                today = datetime.now().strftime("%d-%m-%Y")
                cleaned_name = f"Cleaned_{today}.xlsx"

                output = BytesIO()
                with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                    df_clean.to_excel(writer, index=False, sheet_name='Cleaned Data')
                output.seek(0)

                st.download_button(
                    label="💾 Save Cleaned File",
                    data=output,
                    file_name=cleaned_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"save_{idx}"
                )

                st.caption(f"📁 File will be saved as `{cleaned_name}` in your browser's Downloads folder.")

            except Exception as e:
                st.error(f"❌ Error processing `{file.name}`: {e}")

else:
    st.info("ℹ️ Please upload at least one file to begin processing.")

# 🔷 Sticky Footer
st.markdown(
    """
    <style>
    .footer {
        position: fixed;
        left: 0;
        bottom: 0;
        width: 100%;
        background-color: #f1f1f1;
        color: #333;
        text-align: center;
        padding: 10px;
        font-size: 14px;
        border-top: 1px solid #ccc;
    }
    .reportview-container .main .block-container{
        padding-bottom: 60px; /* space for footer */
    }
    </style>

    <div class="footer">
        <b>Designed By : Sushant Joshi</b>,
        <b>Contact : sushantjoshi800@gmail.com</b>,
        <b>🚀 PunchPro © 2025 — All Rights Reserved</b>
    </div>
    """,
    unsafe_allow_html=True
)