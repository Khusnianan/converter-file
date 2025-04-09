import streamlit as st
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
from docx import Document
import io

st.set_page_config(page_title="PDF to Word Converter", layout="centered")

st.title("📄 PDF to Word Converter")
st.write("Upload file PDF kamu, dan kami akan ubah jadi Word (.docx)")

uploaded_file = st.file_uploader("Upload PDF file", type="pdf")
use_ocr = st.checkbox("Gunakan OCR (untuk PDF hasil scan/gambar)")

if uploaded_file and st.button("Convert to Word"):
    doc = Document()

    if use_ocr:
        st.info("Menggunakan OCR...")
        images = convert_from_bytes(uploaded_file.read())
        for i, img in enumerate(images):
            text = pytesseract.image_to_string(img)
            doc.add_paragraph(text)
            st.success(f"OCR halaman {i+1} selesai")
    else:
        st.info("Ekstrak teks dari PDF...")
        uploaded_file.seek(0)
        with pdfplumber.open(uploaded_file) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    doc.add_paragraph(text)
                st.success(f"Ekstraksi halaman {i+1} selesai")

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.download_button(
        label="📥 Download Word File",
        data=buffer,
        file_name="hasil_konversi.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
