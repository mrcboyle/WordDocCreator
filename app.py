import os,re
from io import BytesIO
from zipfile import ZipFile
import pandas as pd
import streamlit as st
from docxtpl import DocxTemplate
try:
    from jinja2 import meta, Environment
except Exception:
    meta=None

st.set_page_config(page_title="Word Doc Generator",page_icon="📄")
st.title("📄 Word Doc Generator")

@st.cache_data
def load_excel(file,sheet):
    df=pd.read_excel(file,sheet_name=sheet,engine="openpyxl")
    return df.dropna(how="all")

def unique(name,used):
    b,e=os.path.splitext(name);i=1;n=name
    while n in used:
        n=f"{b}_{i}{e}";i+=1
    used.add(n);return n

def clean_filename(s,idx):
    s="" if pd.isna(s) else str(s).strip()
    if not s: s=f"Document_{idx}"
    s=re.sub(r'[\\/*?:"<>|]','',s).replace(" ","_")
    return s

def placeholders(docfile):
    if meta is None: return set()
    z=ZipFile(docfile)
    xml=z.read("word/document.xml").decode("utf-8","ignore")
    env=Environment()
    ast=env.parse(xml)
    return meta.find_undeclared_variables(ast)

excel=st.file_uploader("Excel",type="xlsx")
template=st.file_uploader("Template",type="docx")
sheet=st.text_input("Sheet",value="Sheet1")
zipname=st.text_input("ZIP filename",value="generated_documents.zip")

if excel and template:
    try:
        df=load_excel(excel,sheet)
    except Exception as e:
        st.error(e);st.stop()
    st.dataframe(df.head())
    if "Filename" not in df.columns:
        st.error("Spreadsheet requires a Filename column.");st.stop()
    ph=placeholders(template)
    if ph:
        missing=sorted(ph-set(df.columns))
        if missing:
            st.warning("Template fields missing from spreadsheet: "+", ".join(missing))
    if st.button("Generate Documents"):
        prog=st.progress(0)
        report=[];failed=0;made=0
        used=set();zipbuf=BytesIO()
        with st.spinner("Generating..."):
            with ZipFile(zipbuf,"w") as z:
                for i,(_,row) in enumerate(df.iterrows(),start=1):
                    try:
                        ctx={k:("" if pd.isna(v) else str(v).strip()) for k,v in row.items()}
                        doc=DocxTemplate(template)
                        doc.render(ctx)
                        fname=unique(clean_filename(row["Filename"],i)+".docx",used)
                        bio=BytesIO();doc.save(bio);bio.seek(0)
                        z.writestr(fname,bio.read())
                        made+=1
                        report.append(f"OK: {fname}")
                    except Exception as e:
                        failed+=1
                        report.append(f"FAILED row {i+1}: {e}")
                    prog.progress(i/len(df))
                z.writestr("Generation_Report.txt",
                           f"Generated: {made}\nFailed: {failed}\n\n"+"\n".join(report))
        zipbuf.seek(0)
        st.success(f"Generated {made} documents. Failed: {failed}")
        st.download_button("⬇️ Download ZIP",zipbuf,file_name=zipname,mime="application/zip")