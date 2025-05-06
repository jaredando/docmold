import streamlit as st
from docx import Document
from io import BytesIO

def strip_docx_content(docx_file):
    # Load the uploaded .docx file
    doc = Document(docx_file)
    
    # Replace paragraph text with "[Blank]"
    for para in doc.paragraphs:
        if para.text.strip():  # Skip truly empty paragraphs
            for run in para.runs:
                run.text = "[Blank]"

    # Replace table cell content with "[Blank]"
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if para.text.strip():
                        for run in para.runs:
                            run.text = "[Blank]"

    return doc

def save_docx_to_bytes(doc):
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# Streamlit UI
st.title("DocMold - Strip and Preserve Layout")

uploaded_file = st.file_uploader("Upload a DOCX template", type=["docx"])

if uploaded_file:
    st.success("File uploaded successfully!")

    # Process file
    stripped_doc = strip_docx_content(uploaded_file)
    output = save_docx_to_bytes(stripped_doc)

    # Download button
    st.download_button(
        label="Download Stripped Template",
        data=output,
        file_name="stripped_template.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
