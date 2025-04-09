import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from docx import Document
import pytesseract
from PIL import Image
import os
import tempfile
from io import BytesIO

st.set_page_config(page_title="PDF/Gambar ke Word Converter", layout="centered")

st.title("📄 PDF/Gambar ke Word Converter (dengan Preview & Seleksi Teks)")

menu = st.radio("Pilih tipe file yang ingin dikonversi:", ["PDF", "Gambar (OCR)"])

def display_docx_content(doc_buffer):
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError
    try:
        doc = Document(doc_buffer)
        st.markdown("### 📄 Pratinjau Isi Dokumen Word")
        for para in doc.paragraphs:
            if para.text.strip():
                st.write(para.text)
    except PackageNotFoundError:
        st.error("Gagal membuka dokumen Word untuk pratinjau.")

if menu == "PDF":
    uploaded_file = st.file_uploader("Unggah file PDF", type=["pdf"])

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
                    doc = Document()
                    for para in selected_text:
                        doc.add_paragraph(para)

                    preview_buffer = BytesIO()
                    doc.save(preview_buffer)
                    preview_buffer.seek(0)

                    output_name = os.path.splitext(uploaded_file.name)[0] + " (konversi).docx"
                    st.download_button("⬇️ Unduh Hasil Word", data=preview_buffer, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                    # Tampilkan pratinjau isi Word
                    display_docx_content(preview_buffer)
            else:
                st.warning("Silakan pilih minimal satu halaman.")

elif menu == "Gambar (OCR)":
    uploaded_image = st.file_uploader("Unggah gambar (JPG, PNG)", type=["png", "jpg", "jpeg"])

    if uploaded_image:
        image = Image.open(uploaded_image)
        st.image(image, caption="Gambar yang diunggah", use_column_width=True)

        text = pytesseract.image_to_string(image)

        st.markdown("### 🔍 Pratinjau Teks Hasil OCR")
        paragraphs = text.split("\n")
        selected_text = []

        for i, para in enumerate(paragraphs):
            if para.strip():
                if st.checkbox(para, key=f"ocr-{i}"):
                    selected_text.append(para)

        if selected_text:
            doc = Document()
            for para in selected_text:
                doc.add_paragraph(para)

            preview_buffer = BytesIO()
            doc.save(preview_buffer)
            preview_buffer.seek(0)

            output_name = os.path.splitext(uploaded_image.name)[0] + " (konversi).docx"
            st.download_button("⬇️ Unduh Hasil Word", data=preview_buffer, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            # Tampilkan pratinjau isi Word
            display_docx_content(preview_buffer)
