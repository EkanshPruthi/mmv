import streamlit as st
import pandas as pd
import zipfile
import os
import shutil
from pathlib import Path

# Function to process files
def process_files(excel_file, zip_file):
    # Step 1: Read Excel File
    excel_data = pd.ExcelFile(excel_file)
    sheet = excel_data.parse(excel_data.sheet_names[0])  # Read the first sheet
    st.write("Preview of Excel File:")
    st.dataframe(sheet.head())

    if 'Label' not in sheet.columns or 'District' not in sheet.columns:
        st.error("The Excel file must contain 'Label' and 'District' columns.")
        return

    # Step 2: Unzip the uploaded ZIP file
    output_dir = "temp_unzipped"
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)
    with zipfile.ZipFile(zip_file, 'r') as zip_ref:
        zip_ref.extractall(output_dir)

    st.write("ZIP file extracted successfully!")

    # Step 3: Organize PDFs by district
    organized_dir = "organized_pdfs"
    if os.path.exists(organized_dir):
        shutil.rmtree(organized_dir)

    os.makedirs(organized_dir, exist_ok=True)
    pdf_folder = Path(output_dir)
    pdf_files = {f.name: f for f in pdf_folder.glob("*.pdf")}

    for _, row in sheet.iterrows():
        label = row['Label']
        district = row['District']
        pdf_name = f"{label}.pdf"

        if pdf_name in pdf_files:
            district_dir = Path(organized_dir) / district
            district_dir.mkdir(parents=True, exist_ok=True)
            shutil.copy(pdf_files[pdf_name], district_dir / pdf_name)

    # Step 4: Zip organized folder
    zip_output = "organized_pdfs.zip"
    if os.path.exists(zip_output):
        os.remove(zip_output)

    shutil.make_archive("organized_pdfs", 'zip', organized_dir)

    return zip_output

# Streamlit UI
st.title("Organize PDFs by District")
st.write("Upload an Excel file and a ZIP file containing PDFs to organize them by district.")

# File upload widgets
uploaded_excel = st.file_uploader("Upload Excel File", type=['xlsx'])
uploaded_zip = st.file_uploader("Upload ZIP File with PDFs", type=['zip'])

if uploaded_excel and uploaded_zip:
    # Process the files
    with st.spinner("Processing..."):
        output_zip = process_files(uploaded_excel, uploaded_zip)

    # Download organized zip file
    with open(output_zip, "rb") as f:
        st.download_button(
            label="Download Organized PDFs",
            data=f,
            file_name="organized_pdfs.zip",
            mime="application/zip"
        )
