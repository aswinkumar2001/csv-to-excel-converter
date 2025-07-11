import streamlit as st
import pandas as pd
import zipfile
import os
import io
import shutil
from datetime import datetime

# Streamlit app title
st.title("CSV Folder to Excel Converter")

# File uploader for ZIP file
uploaded_file = st.file_uploader("Upload a ZIP file containing a folder with CSV files", type="zip")

if uploaded_file is not None:
    # Create a temporary directory to extract files
    temp_dir = "temp_extracted"
    if os.path.exists(temp_dir):
        shutil.rmtree(temp_dir)
    os.makedirs(temp_dir)

    # Save and extract the ZIP file
    zip_path = os.path.join(temp_dir, "uploaded.zip")
    with open(zip_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    try:
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
    except zipfile.BadZipFile:
        st.error("Invalid ZIP file. Please upload a valid ZIP file.")
        shutil.rmtree(temp_dir)
        st.stop()

    # Find all CSV files recursively, ignoring hidden files
    csv_files = []
    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.lower().endswith('.csv') and not file.startswith('.'):
                csv_files.append(os.path.join(root, file))
    
    if not csv_files:
        st.error("No CSV files found in the ZIP file. Ensure the ZIP contains CSV files.")
        st.write("Extracted folder contents:", os.listdir(temp_dir))
        for root, dirs, files in os.walk(temp_dir):
            st.write(f"Directory: {root}")
            st.write(f"Subdirectories: {dirs}")
            st.write(f"Files: {files}")
        shutil.rmtree(temp_dir)
        st.stop()
    
    # Create an in-memory Excel file
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for csv_path in csv_files:
            try:
                # Read CSV with UTF-8-SIG encoding to handle BOM
                df = pd.read_csv(csv_path, encoding='utf-8-sig')
                # Get sheet name from CSV file (without .csv extension, max 31 chars for Excel)
                sheet_name = os.path.splitext(os.path.basename(csv_path))[0][:31]
                # Write to Excel sheet
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                st.write(f"Processed: {os.path.basename(csv_path)}")
            except Exception as e:
                st.warning(f"Error processing {os.path.basename(csv_path)}: {str(e)}")
    
    # Prepare file for download
    output.seek(0)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_filename = f"converted_excel_{timestamp}.xlsx"
    
    st.success(f"Converted {len(csv_files)} CSV files to Excel sheets.")
    st.download_button(
        label="Download Excel File",
        data=output,
        file_name=excel_filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    # Clean up temporary directory
    shutil.rmtree(temp_dir)
