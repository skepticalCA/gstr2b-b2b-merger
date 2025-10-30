import streamlit as st
import pandas as pd
import io

# --- Page Configuration ---
st.set_page_config(
    page_title="GSTR 2B Excel Merger",
    page_icon="⚙️",
    layout="centered"
)

# --- Main Application ---
st.title("⚙️ Robust Excel File Merger")
st.write("""
Upload your Excel files. This tool will intelligently find and merge all data sheets from all files, skipping any corrupted sheets and providing a detailed summary of its actions.
""")

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Choose your .xlsx files",
    accept_multiple_files=True,
    type="xlsx"
)

# --- Merge Button and Logic ---
if st.button("Merge Files"):
    if uploaded_files:
        all_data_frames = []
        
        # --- CONFIGURATION BASED ON YOUR CHOICES ---
        
        # 1. Flexible Sheet Name Matching: Normalizing to lowercase and stripping spaces.
        sheets_to_exclude_raw = ["Read me", "ITC Available", "ITC not available", "ITC Reversal", "ITC Rejected"]
        sheets_to_exclude = [name.strip().lower() for name in sheets_to_exclude_raw]
        
        # 3. Detailed Logging: List to hold summary messages for the user.
        process_log = []

        st.write("Starting merge process...")
        progress_bar = st.progress(0)
        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):
            sheets_merged_from_this_file = []
            errors_in_this_file = []

            try:
                # Open the Excel file to inspect its sheets
                xls = pd.ExcelFile(file)
                sheet_names = xls.sheet_names
                
                for sheet_name in sheet_names:
                    normalized_sheet_name = sheet_name.strip().lower()
                    
                    # Check if the normalized sheet name is in our exclusion list
                    if normalized_sheet_name not in sheets_to_exclude:
                        try:
                            # 2. Robust Error Handling: Try to read each sheet individually.
                            df = pd.read_excel(file, sheet_name=sheet_name, skiprows=4, header=[0, 1])
                            
                            # Add trace columns
                            df['Original_Filename'] = file.name
                            df['Original_Sheet_Name'] = sheet_name
                            
                            all_data_frames.append(df)
                            sheets_merged_from_this_file.append(sheet_name)

                        except Exception as e:
                            # If a single sheet fails, log it and continue.
                            errors_in_this_file.append(f"Skipped corrupted/invalid sheet '{sheet_name}' (Error: {e})")

            except Exception as e:
                # Handle errors at the file level (e.g., password protected)
                process_log.append(f"🔴 **{file.name}:** Failed to process entire file. It might be corrupted or password-protected. (Error: {e})")

            # --- LOGGING FOR THE CURRENT FILE ---
            if sheets_merged_from_this_file:
                log_entry = f"🟢 **{file.name}:** Merged sheets: `" + "`, `".join(sheets_merged_from_this_file) + "`"
                if errors_in_this_file:
                    log_entry += " (and skipped some invalid sheets)"
                process_log.append(log_entry)
            elif errors_in_this_file and not sheets_merged_from_this_file:
                 process_log.append(f"🟡 **{file.name}:** Skipped. Contained only invalid or summary sheets.")
            else:
                process_log.append(f"⚪️ **{file.name}:** Skipped (no data sheets found).")

            progress_bar.progress((i + 1) / total_files)
        
        # --- FINAL STEP: CHECK IF ANY DATA WAS MERGED ---
        
        # 4. Handle Empty Results: Only proceed if dataframes were created.
        if all_data_frames:
            st.success(f"Merge complete! Successfully processed {len(uploaded_files)} files.")
            
            # Display the detailed process log
            with st.expander("Click to see the detailed processing log"):
                for entry in process_log:
                    st.markdown(entry)
            
            # Concatenate, flatten headers, and prepare for download
            merged_df = pd.concat(all_data_frames, ignore_index=True)

            if isinstance(merged_df.columns, pd.MultiIndex):
                merged_df.columns = ['_'.join(map(str, col)).strip() for col in merged_df.columns.values]
                merged_df.columns = [col.replace('_Unnamed: 1_level_1', '').replace('Unnamed: 0_level_0_', '') for col in merged_df.columns]

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Consolidated_Data')
            
            processed_data = output.getvalue()

            st.download_button(
                label="📥 Download Consolidated Excel File",
                data=processed_data,
                file_name="consolidated_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            # Display a warning if no data could be merged at all.
            st.warning("No data sheets found to merge across all uploaded files. Please check your files.")
            with st.expander("Click to see the detailed processing log"):
                for entry in process_log:
                    st.markdown(entry)
    else:
        st.warning("Please upload at least one Excel file.")
