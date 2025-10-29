import streamlit as st
import pandas as pd
import os
from io import BytesIO
from zipfile import ZipFile
from docxtpl import DocxTemplate

st.title("📄 Word Doc Generator")

# === Upload Inputs ===
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Word Template", type=["docx"])
sheet_name = st.text_input("Sheet Name", value="Sheet1")

# === Helper Function ===
def get_unique_filename(filename, existing_names):
    base, ext = os.path.splitext(filename)
    counter = 1
    while filename in existing_names:
        filename = f"{base}_{counter}{ext}"
        counter += 1
    return filename

# === Main Logic ===
if excel_file and template_file:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
    df = df.dropna(how='all')
    st.write("📊 Preview of Data", df.head())

    if "Filename" not in df.columns:
        st.error("❌ The spreadsheet must contain a 'Filename' column for output filenames.")
    elif st.button("Generate Documents"):
        zip_buffer = BytesIO()
        existing_names = set()

        with ZipFile(zip_buffer, "w") as zip_file:
            for _, row in df.iterrows():
                context = row.to_dict()
                doc = DocxTemplate(template_file)
                doc.render(context)

                raw_name = str(row["Filename"]).strip().replace(" ", "_")
                filename = get_unique_filename(f"{raw_name}.docx", existing_names)
                existing_names.add(filename)

                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                zip_file.writestr(filename, doc_io.read())

        zip_buffer.seek(0)

        st.download_button(
            label="⬇️ Download All Documents as ZIP",
            data=zip_buffer,
            file_name="generated_documents.zip",
            mime="application/zip"
        )

        st.success("✅ All documents created and zipped successfully.")