import streamlit as st
import fitz  # PyMuPDF
import pdfplumber
from docx import Document
import pytesseract
from PIL import Image
import os
import tempfile
from io import BytesIO
from docx2pdf import convert
import base64

st.set_page_config(page_title="PDF/Gambar ke Word Converter", layout="centered")

st.title("📄 PDF/Gambar ke Word Converter (dengan Preview Word dan Seleksi)")

menu = st.radio("Pilih tipe file yang ingin dikonversi:", ["PDF", "Gambar (OCR)"])

def show_pdf(file_path):
    with open(file_path, "rb") as f:
        base64_pdf = base64.b64encode(f.read()).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)

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

                    docx_buffer = BytesIO()
                    doc.save(docx_buffer)
                    docx_buffer.seek(0)

                    # Save to temp file to convert to PDF for preview
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                        tmp_docx.write(docx_buffer.getvalue())
                        tmp_docx_path = tmp_docx.name

                    tmp_pdf_preview_path = tmp_docx_path.replace(".docx", ".pdf")
                    convert(tmp_docx_path, tmp_pdf_preview_path)

                    output_name = os.path.splitext(uploaded_file.name)[0] + " (konversi).docx"
                    st.download_button("⬇️ Unduh Hasil Word", data=docx_buffer, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

                    st.markdown("### 📄 Pratinjau Dokumen Word (dalam bentuk PDF)")
                    show_pdf(tmp_pdf_preview_path)
                else:
                    st.warning("Tidak ada teks yang dipilih untuk dikonversi.")
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

            docx_buffer = BytesIO()
            doc.save(docx_buffer)
            docx_buffer.seek(0)

            # Save to temp file to convert to PDF for preview
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp_docx:
                tmp_docx.write(docx_buffer.getvalue())
                tmp_docx_path = tmp_docx.name

            tmp_pdf_preview_path = tmp_docx_path.replace(".docx", ".pdf")
            convert(tmp_docx_path, tmp_pdf_preview_path)

            output_name = os.path.splitext(uploaded_image.name)[0] + " (konversi).docx"
            st.download_button("⬇️ Unduh Hasil Word", data=docx_buffer, file_name=output_name, mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

            st.markdown("### 📄 Pratinjau Dokumen Word (dalam bentuk PDF)")
            show_pdf(tmp_pdf_preview_path)
        else:
            st.warning("Tidak ada teks yang dipilih untuk dikonversi.")
