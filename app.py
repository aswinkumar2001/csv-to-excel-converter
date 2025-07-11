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
    
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(temp_dir)
    
    # Find the extracted folder (assuming ZIP contains a single folder)
    extracted_folders = [f for f in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, f))]
    if not extracted_folders:
        st.error("The ZIP file does not contain a folder. Please upload a ZIP with a folder containing CSV files.")
    else:
        csv_folder = os.path.join(temp_dir, extracted_folders[0])
        
        # Get list of CSV files
        csv_files = [f for f in os.listdir(csv_folder) if f.lower().endswith('.csv')]
        
        if not csv_files:
            st.error("No CSV files found in the folder.")
        else:
            # Create an in-memory Excel file
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                for csv_file in csv_files:
                    # Read CSV
                    csv_path = os.path.join(csv_folder, csv_file)
                    try:
                        df = pd.read_csv(csv_path)
                        # Get sheet name from CSV file (without .csv extension)
                        sheet_name = os.path.splitext(csv_file)[0]
                        # Write to Excel sheet
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                    except Exception as e:
                        st.warning(f"Error processing {csv_file}: {str(e)}")
            
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