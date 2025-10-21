import streamlit as st
import pandas as pd
import os
from io import BytesIO
from zipfile import ZipFile
from docxtpl import DocxTemplate

st.title("📄 Chris B's Marking Review Form Generator")

# === Upload Inputs ===
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Word Template", type=["docx"])
sheet_name = st.text_input("Sheet Name", value="CombinedSubjects")

# === Helper Function ===
def get_unique_path(path, existing_paths):
    base, ext = os.path.splitext(path)
    counter = 1
    while path in existing_paths:
        path = f"{base}_{counter}{ext}"
        counter += 1
    return path

# === Main Logic ===
if excel_file and template_file:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
    df = df.dropna(how='all')
    st.write("📊 Preview of Data", df.head())

    if st.button("Generate Documents"):
        zip_buffer = BytesIO()
        existing_paths = set()

        with ZipFile(zip_buffer, "w") as zip_file:
            for idx, row in df.iterrows():
                context = row.to_dict()
                doc = DocxTemplate(template_file)
                doc.render(context)

                subject = str(row['Subject']).strip().replace(" ", "_")
                assessment_code = str(row['Assessment_Code']).strip().replace(" ", "_")
                filename = f"{row['Assessment_Code']}-{row['Assessment_Type']}.docx".replace(" ", "_")

                zip_path = os.path.join(subject, assessment_code, filename)
                zip_path = get_unique_path(zip_path, existing_paths)
                existing_paths.add(zip_path)

                doc_io = BytesIO()
                doc.save(doc_io)
                doc_io.seek(0)

                zip_file.writestr(zip_path, doc_io.read())

        zip_buffer.seek(0)

        st.download_button(
            label="⬇️ Download All Documents as ZIP",
            data=zip_buffer,
            file_name="generated_documents.zip",
            mime="application/zip"
        )

        st.success("✅ All documents created and zipped with folder structure.")