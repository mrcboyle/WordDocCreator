import streamlit as st
import pandas as pd
import os
from docxtpl import DocxTemplate

st.title("📄 Marking Review Form Generator")

# === Upload Inputs ===
excel_file = st.file_uploader("Upload Excel File", type=["xlsx"])
template_file = st.file_uploader("Upload Word Template", type=["docx"])
output_base = st.text_input("Output Folder Path")

# === Sheet Selection ===
sheet_name = st.text_input("Sheet Name", value="CombinedSubjects")

def get_unique_path(path):
    base, ext = os.path.splitext(path)
    counter = 1
    while os.path.exists(path):
        path = f"{base}_{counter}{ext}"
        counter += 1
    return path

if excel_file and template_file and output_base:
    df = pd.read_excel(excel_file, sheet_name=sheet_name, engine='openpyxl')
    df = df.dropna(how='all')
    st.write("📊 Preview of Data", df.head())

    if st.button("Generate Documents"):
        for idx, row in df.iterrows():
            context = row.to_dict()
            doc = DocxTemplate(template_file)
            doc.render(context)

            subject = str(row['Subject']).strip().replace(" ", "_")
            assessment_code = str(row['Assessment_Code']).strip().replace(" ", "_")
            output_folder = os.path.join(output_base, subject, assessment_code)
            os.makedirs(output_folder, exist_ok=True)

            filename = f"{row['Assessment_Code']}-{row['Assessment_Type']}.docx".replace(" ", "_")
            full_path = os.path.join(output_folder, filename)
            full_path = get_unique_path(full_path)

            doc.save(full_path)
            st.success(f"✔️ Saved: {full_path}")

        st.info("✅ All documents created with formatting and images intact.")
