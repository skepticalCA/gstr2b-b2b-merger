import streamlit as st
import pandas as pd
import io

# --- Page Configuration ---
st.set_page_config(
    page_title="Excel File Merger",
    page_icon="📄",
    layout="centered"
)

# --- Main Application ---
st.title("📄 Excel File Merger")
st.write("Upload your GSTR-2B Excel files, and this tool will merge the 'B2B' sheets into a single file for you to download.")

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
        st.write("Processing files...")

        progress_bar = st.progress(0)
        total_files = len(uploaded_files)

        for i, file in enumerate(uploaded_files):
            try:
                # Read the 'B2B' sheet from the uploaded file
                df = pd.read_excel(file, sheet_name='B2B', skiprows=4, header=[0, 1])
                df['Original_Filename'] = file.name
                all_data_frames.append(df)

            except Exception as e:
                st.error(f"Error processing '{file.name}': {e}. Please check the file format and ensure it has a 'B2B' sheet.")

            progress_bar.progress((i + 1) / total_files)

        if all_data_frames:
            # Concatenate all dataframes
            merged_df = pd.concat(all_data_frames, ignore_index=True)

            # Flatten the multi-level headers
            if isinstance(merged_df.columns, pd.MultiIndex):
                merged_df.columns = ['_'.join(col).strip() for col in merged_df.columns.values]
                merged_df.columns = [col.replace('Unnamed: 0_level_0_', '') if 'Unnamed' in col else col for col in merged_df.columns]
            
            st.success(f"Successfully merged {len(all_data_frames)} files!")

            # Convert dataframe to Excel in memory
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                merged_df.to_excel(writer, index=False, sheet_name='Merged_B2B_Data')
            
            processed_data = output.getvalue()

            # --- Download Button ---
            st.download_button(
                label="📥 Download Merged Excel File",
                data=processed_data,
                file_name="merged_b2b_data.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    else:
        st.warning("Please upload at least one Excel file.")