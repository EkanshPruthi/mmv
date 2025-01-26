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
def organize_files(excel_file, zip_file):
    # Step 1: Save uploaded Excel file
    with open("uploaded_excel.xlsx", "wb") as f:
        f.write(excel_file.getbuffer())
    
    # Step 2: Read the Excel file
    data = pd.read_excel("uploaded_excel.xlsx")
    required_columns = ["State Name", "District Name", "Label Number 1", "Label Number 2", "Label Number 3", "Label Number 4"]

    # Validate Excel columns
    for col in required_columns:
        if col not in data.columns:
            st.error(f"Missing required column: {col}")
            return None

    # Step 3: Extract PDFs from the uploaded ZIP file
    temp_pdf_folder = "temp_pdfs"
    ensure_directory(temp_pdf_folder)

    # Save and extract the uploaded ZIP file
    zip_file_path = "uploaded_pdfs.zip"
    with open(zip_file_path, "wb") as f:
        f.write(zip_file.getbuffer())

    # Log ZIP contents and extract only PDF files
    with ZipFile(zip_file_path, 'r') as zip_ref:
        st.write("Files in ZIP:", zip_ref.namelist())  # Log all files in the ZIP
        for file in zip_ref.namelist():
            if file.endswith(".pdf"):  # Only process .pdf files
                file_name = os.path.basename(file)  # Extract file name
                with open(os.path.join(temp_pdf_folder, file_name), "wb") as f:
                    f.write(zip_ref.read(file))

    # Log extracted files for debugging
    extracted_files = [f.name for f in Path(temp_pdf_folder).glob("*.pdf")]
    st.write("Extracted Files:", extracted_files)

    # Step 4: Organize PDFs into folders
    organized_folder = "Organized_Files"
    ensure_directory(organized_folder)
    pdf_files = {f.name: f for f in Path(temp_pdf_folder).glob("*.pdf")}  # Case-sensitive match

    st.write("Labels in Excel:", data[["Label Number 1", "Label Number 2", "Label Number 3", "Label Number 4"]].dropna().values.flatten())

    for _, row in data.iterrows():
        state = row["State Name"]
        district = row["District Name"]
        label_columns = ["Label Number 1", "Label Number 2", "Label Number 3", "Label Number 4"]

        # Create state and district folders
        state_folder = Path(organized_folder) / state
        district_folder = state_folder / district
        district_folder.mkdir(parents=True, exist_ok=True)

        for col in label_columns:
            if pd.notna(row[col]):  # Check if the label exists
                pdf_name = row[col].strip()  # Trim spaces
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
st.write("Upload an Excel file and a ZIP file containing PDFs to organize them by state and district.")

# File upload widgets
uploaded_excel = st.file_uploader("Upload Excel File", type=["xlsx"])
uploaded_zip = st.file_uploader("Upload ZIP File of PDFs", type=["zip"])

if uploaded_excel and uploaded_zip:
    with st.spinner("Processing files..."):
        output_zip = organize_files(uploaded_excel, uploaded_zip)

        if output_zip:
            st.success("Files organized successfully!")
            with open(output_zip, "rb") as f:
                st.download_button(
                    label="Download Organized PDFs",
                    data=f,
                    file_name="Organized_Files.zip",
                    mime="application/zip"
                )
