import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from docx import Document
import os
import tempfile

st.set_page_config(page_title="PDF ke Word Converter", layout="centered")

st.title("📄 PDF ke Word Converter (Preview + Pilih Halaman/Teks)")

uploaded_file = st.file_uploader("Unggah file PDF kamu di sini", type=["pdf"])

if uploaded_file:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
        tmp_file.write(uploaded_file.read())
        tmp_pdf_path = tmp_file.name

    with pdfplumber.open(tmp_pdf_path) as pdf:
        total_pages = len(pdf.pages)
        st.info(f"📄 File memiliki {total_pages} halaman.")

        selected_pages = st.multiselect(
            "Pilih halaman yang ingin dikonversi:",
            options=list(range(1, total_pages + 1)),
            default=list(range(1, total_pages + 1)),
        )

        if selected_pages:
            all_extracted_text = []
            for page_num in selected_pages:
                page = pdf.pages[page_num - 1]
                text = page.extract_text()
                if text:
                    all_extracted_text.append((page_num, text.strip()))
                else:
                    all_extracted_text.append((page_num, "[Halaman kosong atau tidak bisa dibaca]"))

            st.markdown("### 🔍 Pratinjau & Pilih Teks")
            selected_text = []

            for page_num, text in all_extracted_text:
                with st.expander(f"Halaman {page_num}"):
                    paragraphs = text.split("\n")
                    for i, para in enumerate(paragraphs):
                        if st.checkbox(f"[Hal. {page_num}] {para}", key=f"{page_num}-{i}"):
                            selected_text.append(para)

            if selected_text:
                if st.button("📥 Konversi dan Unduh Word"):
                    doc = Document()
                    for para in selected_text:
                        doc.add_paragraph(para)

                    output_name = os.path.splitext(uploaded_file.name)[0] + " (konversi).docx"
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                        doc.save(tmp_docx.name)
                        with open(tmp_docx.name, "rb") as f:
                            st.download_button("⬇️ Unduh Hasil Word", data=f, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        else:
            st.warning("Silakan pilih minimal satu halaman.")
else:
    st.info("Silakan unggah file PDF terlebih dahulu.")
