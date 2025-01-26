import streamlit as st
import pandas as pd
import os
import shutil
from pathlib import Path
from zipfile import ZipFile

# Ensure directory exists
def ensure_directory(directory):
    if not os.path.exists(directory):
        os.makedirs(directory)

# Organize files into state-wise and district-wise folders
def organize_files(excel_file, uploaded_pdfs):
    # Step 1: Save uploaded Excel file
    with open("uploaded_excel.xlsx", "wb") as f:
        f.write(excel_file.read())
    
    # Step 2: Read the Excel file
    data = pd.read_excel("uploaded_excel.xlsx")
    required_columns = ["State Name", "District Name", "Label Number 1", "Label Number 2", "Label Number 3", "Label Number 4"]

    # Validate Excel columns
    for col in required_columns:
        if col not in data.columns:
            st.error(f"Missing required column: {col}")
            return None

    # Step 3: Save uploaded PDFs
    pdf_folder = "temp_pdfs"
    ensure_directory(pdf_folder)
    for uploaded_file in uploaded_pdfs:
        with open(os.path.join(pdf_folder, uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())

    # Step 4: Organize PDFs into folders
    organized_folder = "Organized_Files"
    ensure_directory(organized_folder)
    pdf_files = {f.name: f for f in Path(pdf_folder).glob("*.pdf")}

    for _, row in data.iterrows():
        state = row["State Name"]
        district = row["District Name"]
        label_columns = ["Label Number 1", "Label Number 2", "Label Number 3", "Label Number 4"]

        # Create state and district folders
        state_folder = Path(organized_folder) / state
        district_folder = state_folder / district
        district_folder.mkdir(parents=True, exist_ok=True)

        for col in label_columns:
            if pd.notna(row[col]):
                pdf_name = f"{row[col]}.pdf"
                if pdf_name in pdf_files:
                    shutil.copy(pdf_files[pdf_name], district_folder / pdf_name)
                else:
                    st.warning(f"File not found: {pdf_name}")

    # Step 5: Zip the organized folder
    zip_output = "Organized_Files.zip"
    with ZipFile(zip_output, "w") as zipf:
        for root, _, files in os.walk(organized_folder):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, organized_folder)
                zipf.write(file_path, arcname)

    return zip_output

# Streamlit UI
st.title("Organize PDFs by State and District")
st.write("Upload an Excel file and multiple PDF files to organize them by state and district.")

# File upload widgets
uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx"])
uploaded_pdfs = st.file_uploader("Upload PDF Files", type=["pdf"], accept_multiple_files=True)

if uploaded_excel and uploaded_pdfs:
    with st.spinner("Processing files..."):
        output_zip = organize_files(uploaded_excel, uploaded_pdfs)

        if output_zip:
            st.success("Files organized successfully!")
            with open(output_zip, "rb") as f:
                st.download_button(
                    label="Download Organized PDFs",
                    data=f,
                    file_name="Organized_Files.zip",
                    mime="application/zip"
                )
