
import os,re
from io import BytesIO
from zipfile import ZipFile
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate

st.set_page_config(page_title="Word Doc Generator", page_icon="📄")
st.title("📄 Word Doc Generator")

@st.cache_data
def load_excel(file,sheet):
    df=pd.read_excel(file,sheet_name=sheet,engine="openpyxl")
    return df.dropna(how="all")

def get_unique_filename(filename,used):
    base,ext=os.path.splitext(filename)
    name=filename;i=1
    while name in used:
        name=f"{base}_{i}{ext}";i+=1
    used.add(name)
    return name

def clean_filename(value,index):
    if pd.isna(value): value=""
    name=str(value).strip() or f"Document_{index}"
    return re.sub(r'[\\/*?:"<>|]','',name).replace(" ","_")

excel_file=st.file_uploader("Upload Excel file",type=["xlsx"])
template_file=st.file_uploader("Upload Word template",type=["docx"])
sheet_name=st.text_input("Worksheet name",value="Sheet1")
zip_filename=st.text_input("ZIP filename",value="generated_documents.zip")

if excel_file and template_file:
    try:
        df=load_excel(excel_file,sheet_name)
    except Exception as e:
        st.error(f"Unable to read Excel file: {e}")
        st.stop()
    st.subheader("Preview")
    st.dataframe(df.head())
    if "Filename" not in df.columns:
        st.error("Spreadsheet must contain a 'Filename' column.")
        st.stop()
    template_bytes=template_file.getvalue()
    if st.button("Generate Documents"):
        progress=st.progress(0)
        zip_buffer=BytesIO()
        used=set();generated=0;failed=0;report=[]
        with st.spinner("Generating documents..."):
            with ZipFile(zip_buffer,"w") as z:
                total=len(df)
                for i,(_,row) in enumerate(df.iterrows(),start=1):
                    try:
                        context={k:("" if pd.isna(v) else str(v).strip()) for k,v in row.items()}
                        doc=DocxTemplate(BytesIO(template_bytes))
                        doc.render(context)
                        fname=get_unique_filename(clean_filename(row["Filename"],i)+".docx",used)
                        b=BytesIO(); doc.save(b); b.seek(0)
                        z.writestr(fname,b.read())
                        generated+=1; report.append(f"SUCCESS: {fname}")
                    except Exception as e:
                        failed+=1; report.append(f"FAILED row {i+1}: {e}")
                    progress.progress(i/total)
                z.writestr("Generation_Report.txt",f"Generated: {generated}\nFailed: {failed}\n\n"+"\n".join(report))
        progress.empty()
        zip_buffer.seek(0)
        st.success(f"Finished! Generated {generated} document(s). Failed: {failed}.")
        st.download_button("⬇️ Download ZIP",data=zip_buffer,file_name=zip_filename,mime="application/zip")