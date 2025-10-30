import streamlit as st
import pandas as pd
import io
import zipfile

# --- Page Configuration ---
st.set_page_config(
    page_title="State-Based Excel Merger",
    page_icon="🗺️",
    layout="centered"
)

# --- Main Application ---
st.title("🗺️ State-Based Excel File Merger")
st.write("""
Upload your GSTR-2B Excel files. This tool identifies the state code from each filename (e.g., '36' from '..._36AAACN...'), merges all periods for each state, and bundles the results into a single downloadable .zip file.
""")

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Choose your .xlsx files",
    accept_multiple_files=True,
    type="xlsx"
)

# --- Merge Button and Logic ---
if st.button("Group and Merge Files by State"):
    if uploaded_files:
        # Use a dictionary to group dataframes by state code
        data_by_state = {}
        
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
                # --- NEW: EXTRACT STATE CODE FROM FILENAME ---
                parts = file.name.split('_')
                if len(parts) > 1 and len(parts[1]) >= 2:
                    state_code = parts[1][:2]
                else:
                    process_log.append(f"🟡 **{file.name}:** Skipped. Filename does not follow the expected format '..._STATECODE...'.")
                    continue # Skip to the next file

                xls = pd.ExcelFile(file)
                sheet_names = xls.sheet_names
                
                for sheet_name in sheet_names:
                    normalized_sheet_name = sheet_name.strip().lower()
                    if normalized_sheet_name not in sheets_to_exclude:
                        try:
                            df = pd.read_excel(file, sheet_name=sheet_name, skiprows=4, header=[0, 1])
                            df['Original_Filename'] = file.name
                            df['Original_Sheet_Name'] = sheet_name
                            
                            # Add the dataframe to the correct state group
                            if state_code not in data_by_state:
                                data_by_state[state_code] = []
                            data_by_state[state_code].append(df)
                            
                            sheets_merged_from_this_file.append(sheet_name)
                        except Exception:
                            # Silently skip corrupted/invalid sheets within a file, as per previous logic
                            pass
                
                if sheets_merged_from_this_file:
                    process_log.append(f"🟢 **{file.name}:** Added sheets `{', '.join(sheets_merged_from_this_file)}` to State Group `{state_code}`.")
                else:
                    process_log.append(f"⚪️ **{file.name}:** Skipped (no data sheets found).")

            except Exception as e:
                process_log.append(f"🔴 **{file.name}:** Failed to process entire file. (Error: {e})")
            
            progress_bar.progress((i + 1) / total_files)
        
        # --- FINAL STEP: PROCESS AND ZIP THE GROUPED DATA ---
        if data_by_state:
            st.success(f"Merge complete! Found data for {len(data_by_state)} different states.")
            
            with st.expander("Click to see the detailed processing log"):
                for entry in process_log:
                    st.markdown(entry)

            # Create an in-memory zip file
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                for state_code, df_list in data_by_state.items():
                    # Merge all dataframes for the current state
                    state_merged_df = pd.concat(df_list, ignore_index=True)

                    # Flatten headers
                    if isinstance(state_merged_df.columns, pd.MultiIndex):
                        state_merged_df.columns = ['_'.join(map(str, col)).strip() for col in state_merged_df.columns.values]
                        state_merged_df.columns = [col.replace('_Unnamed: 1_level_1', '').replace('Unnamed: 0_level_0_', '') for col in state_merged_df.columns]

                    # Create an in-memory Excel file for the current state
                    excel_buffer = io.BytesIO()
                    with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                        state_merged_df.to_excel(writer, index=False, sheet_name=f'State_{state_code}_Data')
                    
                    # Add the in-memory Excel file to the zip archive
                    excel_buffer.seek(0)
                    zipf.writestr(f"State_{state_code}_Consolidated.xlsx", excel_buffer.read())

            st.download_button(
                label="📥 Download All State Files (.zip)",
                data=zip_buffer.getvalue(),
                file_name="Statewise_Consolidated_Data.zip",
                mime="application/zip"
            )
        else:
            st.warning("No data sheets found to merge across all uploaded files. Please check your files and their naming convention.")
            with st.expander("Click to see the detailed processing log"):
                for entry in process_log:
                    st.markdown(entry)
    else:
        st.warning("Please upload at least one Excel file.")
