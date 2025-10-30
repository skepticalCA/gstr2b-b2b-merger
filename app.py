import streamlit as st
import pandas as pd
import io
import zipfile

# --- Page Configuration ---
st.set_page_config(
    page_title="Advanced Excel Merger",
    page_icon="🧩",
    layout="centered"
)

# --- Main Application ---
st.title("🧩 Advanced Excel Merger")
st.write("""
Upload your GSTR-2B files. This tool sorts files by state code, then merges data into separate sheets (B2B, CDNR, etc.) within each state's Excel file. The final result is a single .zip archive.
""")

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Choose your .xlsx files",
    accept_multiple_files=True,
    type="xlsx"
)

# --- Merge Button and Logic ---
if st.button("Group and Merge Files"):
    if uploaded_files:
        # --- NEW: Nested dictionary to group dataframes by state and then by sheet name ---
        # Structure: { 'state_code': { 'sheet_name': [df1, df2, ...], ... }, ... }
        data_grouped = {}
        
        # --- CONFIGURATION ---
        sheets_to_exclude_raw = ["Read me", "ITC Available", "ITC not available", "ITC Reversal", "ITC Rejected"]
        sheets_to_exclude = [name.strip().lower() for name in sheets_to_exclude_raw]
        process_log = []

        st.write("Starting merge process...")
        progress_bar = st.progress(0)
        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):
            sheets_merged_from_this_file = []
            try:
                # Extract state code from filename
                parts = file.name.split('_')
                if len(parts) > 1 and len(parts[1]) >= 2:
                    state_code = parts[1][:2]
                else:
                    process_log.append(f"🟡 **{file.name}:** Skipped. Filename format incorrect.")
                    continue

                xls = pd.ExcelFile(file)
                for sheet_name in xls.sheet_names:
                    normalized_sheet_name = sheet_name.strip().lower()
                    if normalized_sheet_name not in sheets_to_exclude:
                        try:
                            df = pd.read_excel(file, sheet_name=sheet_name, skiprows=4, header=[0, 1])
                            df['Original_Filename'] = file.name
                            
                            # --- NEW GROUPING LOGIC ---
                            # Ensure state_code key exists
                            if state_code not in data_grouped:
                                data_grouped[state_code] = {}
                            # Ensure sheet_name key exists for that state
                            if sheet_name not in data_grouped[state_code]:
                                data_grouped[state_code][sheet_name] = []
                            
                            data_grouped[state_code][sheet_name].append(df)
                            sheets_merged_from_this_file.append(sheet_name)
                        except Exception:
                            pass # Silently skip invalid sheets within a file

                if sheets_merged_from_this_file:
                    process_log.append(f"🟢 **{file.name}:** Added sheets `{', '.join(sheets_merged_from_this_file)}` to State `{state_code}`.")
                else:
                    process_log.append(f"⚪️ **{file.name}:** Skipped (no data sheets found).")
            
            except Exception as e:
                process_log.append(f"🔴 **{file.name}:** Failed to process entire file. (Error: {e})")

            progress_bar.progress((i + 1) / total_files)
        
        # --- FINAL STEP: PROCESS AND ZIP THE GROUPED DATA ---
        if data_grouped:
            st.success(f"Merge complete! Found data for {len(data_grouped)} different states.")
            
            with st.expander("Click to see the detailed processing log"):
                for entry in process_log:
                    st.markdown(entry)

            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                # Iterate through each state in our grouped data
                for state_code, sheets_data in data_grouped.items():
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        # Iterate through each sheet type for the current state
                        for sheet_name, df_list in sheets_data.items():
                            # Merge all dataframes for the current sheet type
                            merged_sheet_df = pd.concat(df_list, ignore_index=True)
                            
                            # Flatten headers
                            if isinstance(merged_sheet_df.columns, pd.MultiIndex):
                                merged_sheet_df.columns = ['_'.join(map(str, col)).strip() for col in merged_sheet_df.columns.values]
                                merged_sheet_df.columns = [col.replace('_Unnamed: 1_level_1', '').replace('Unnamed: 0_level_0_', '') for col in merged_sheet_df.columns]
                            
                            # Write this merged dataframe to a sheet named after its type
                            merged_sheet_df.to_excel(writer, index=False, sheet_name=sheet_name)
                    
                    excel_buffer.seek(0)
                    zipf.writestr(f"State_{state_code}_Consolidated.xlsx", excel_buffer.read())

            st.download_button(
                label="📥 Download All State Files (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Statewise_Consolidated_Data.zip",
                mime="application/zip"
            )
        else:
            st.warning("No data sheets found to merge. Please check your files and naming.")
            with st.expander("Click to see the detailed processing log"):
                for entry in process_log:
                    st.markdown(entry)
    else:
        st.warning("Please upload at least one Excel file.")
