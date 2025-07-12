import streamlit as st
import pandas as pd
from datetime import datetime
import re
from io import BytesIO

# Session state initialization
if "uploaded_files" not in st.session_state:
    st.session_state.uploaded_files = []

st.set_page_config(page_title="PunchPro with Save Button", layout="wide")
st.title("📂 PunchPro")

st.write(
    "Upload one or more Excel/CSV files below. "
    "Each will open in its own tab with cleaned data. "
    "Click **Save Cleaned File** to download in a consistent format."
)

# File uploader (updates session state)
uploaded = st.file_uploader(
    "Select one or more files",
    type=["csv", "xlsx", "xls"],
    accept_multiple_files=True,
    key="file_uploader"
)

# If new upload occurs, update session state
if uploaded:
    st.session_state.uploaded_files = uploaded

# Close button to reset uploaded files
if st.session_state.uploaded_files:
    if st.button("Close", use_container_width=True):
        st.session_state.uploaded_files = []
        try:
            st.experimental_rerun()
        except AttributeError:
            current_params = st.query_params
            new_params = dict(current_params)
            new_params["rerun"] = ["1"] if current_params.get("rerun") != ["1"] else ["0"]
            st.query_params = new_params

# Function to find header row containing 'Punch Records'
def find_header_row(df):
    for i in range(min(10, len(df))):  # Check first 10 rows for header row
        row_values = df.iloc[i].astype(str).str.lower()
        if any('punch records' in val for val in row_values):
            return i
    return None

# Main processing if files are available
if st.session_state.uploaded_files:
    tabs = st.tabs([file.name for file in st.session_state.uploaded_files])

    for idx, file in enumerate(st.session_state.uploaded_files):
        with tabs[idx]:
            st.markdown(f"### 📄 `{file.name}`")

            try:
                # Robust file reading for CSV and Excel formats
                filename = file.name.lower()

                if filename.endswith(".csv"):
                    # Try utf-8, fallback to latin1
                    try:
                        df_raw = pd.read_csv(file, header=None, encoding="utf-8")
                    except UnicodeDecodeError:
                        df_raw = pd.read_csv(file, header=None, encoding="latin1")

                elif filename.endswith(".xls"):
                    # Try 'xlrd' engine (supports old .xls)
                    try:
                        df_raw = pd.read_excel(file, header=None, engine="xlrd")
                    except Exception:
                        # fallback to openpyxl
                        df_raw = pd.read_excel(file, header=None, engine="openpyxl")

                else:
                    # For .xlsx and others, try openpyxl
                    try:
                        df_raw = pd.read_excel(file, header=None, engine="openpyxl")
                    except Exception:
                        # fallback to default engine
                        df_raw = pd.read_excel(file, header=None)

                # Find header row dynamically
                header_row_idx = find_header_row(df_raw)
                if header_row_idx is None:
                    raise ValueError("No header row containing 'Punch Records' found.")

                # Set header and clean data
                df_clean = df_raw.iloc[header_row_idx:].reset_index(drop=True)
                df_clean.columns = df_clean.iloc[0].str.strip()  # strip spaces
                df_clean = df_clean[1:].reset_index(drop=True)

                # Drop irrelevant columns
                if 'S.No' in df_clean.columns:
                    df_clean = df_clean.drop(columns=['S.No'])

                df_clean = df_clean.loc[:, ~df_clean.columns.isna()]

                # Check for Punch Records column again after cleaning
                if 'Punch Records' not in df_clean.columns:
                    st.warning("⚠️ 'Punch Records' column not found. Available columns: " +
                               ", ".join(map(str, df_clean.columns)))
                    raise ValueError("Missing 'Punch Records' column")

                # Extract multiple In/Out times dynamically
                def extract_multiple_in_out(punch_record):
                    if pd.isna(punch_record):
                        return {}

                    entries = re.findall(r'(\d{1,2}:\d{2}:\d{2})\((in|out)\)', punch_record.lower())

                    punch_dict = {}
                    in_count, out_count = 1, 1
                    for time_str, status in entries:
                        if status == "in":
                            punch_dict[f"Time In {in_count}"] = time_str
                            in_count += 1
                        elif status == "out":
                            punch_dict[f"Time Out {out_count}"] = time_str
                            out_count += 1

                    return punch_dict

                # Apply to Punch Records and expand
                expanded_df = df_clean['Punch Records'].apply(extract_multiple_in_out).apply(pd.Series)
                df_clean = pd.concat([df_clean.drop(columns=['Punch Records']), expanded_df], axis=1)

                # Calculate Stay Duration for each pair
                def calculate_stay_durations(row):
                    result = {}
                    i = 1
                    while f"Time In {i}" in row and f"Time Out {i}" in row:
                        time_in = row.get(f"Time In {i}")
                        time_out = row.get(f"Time Out {i}")
                        try:
                            if pd.notna(time_in) and pd.notna(time_out):
                                t_in = pd.to_datetime(time_in, format="%H:%M:%S")
                                t_out = pd.to_datetime(time_out, format="%H:%M:%S")
                                diff = t_out - t_in
                                total_minutes = diff.total_seconds() / 60
                                hours = int(total_minutes // 60)
                                minutes = int(total_minutes % 60)
                                result[f"Stay Duration {i}"] = f"{hours:02}:{minutes:02}"
                            else:
                                result[f"Stay Duration {i}"] = ""
                        except:
                            result[f"Stay Duration {i}"] = ""
                        i += 1
                    return pd.Series(result)

                # Calculate and merge Stay Durations
                stay_durations_df = df_clean.apply(calculate_stay_durations, axis=1)
                df_clean = pd.concat([df_clean, stay_durations_df], axis=1)

                # Reorder columns
                fixed_cols = [col for col in df_clean.columns if not re.match(r'Time (In|Out) \d+|Stay Duration \d+', col)]
                punches = {}
                for col in df_clean.columns:
                    m = re.match(r"(Time In|Time Out|Stay Duration) (\d+)", col)
                    if m:
                        group = int(m.group(2))
                        punches.setdefault(group, {})[m.group(1)] = col

                reordered_punch_cols = []
                for i in sorted(punches.keys()):
                    reordered_punch_cols += [
                        punches[i].get("Time In", f"Time In {i}"),
                        punches[i].get("Time Out", f"Time Out {i}"),
                        punches[i].get("Stay Duration", f"Stay Duration {i}")
                    ]

                final_cols = fixed_cols + reordered_punch_cols
                df_clean = df_clean[[col for col in final_cols if col in df_clean.columns]]

                # Show cleaned data
                st.dataframe(df_clean, use_container_width=True)

                # Save button
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

# Footer
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
        padding-bottom: 60px;
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
