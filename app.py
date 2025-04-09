import streamlit as st
import pdfplumber
import pytesseract
from pdf2image import convert_from_bytes
from docx import Document
import io
from PIL import Image
import re

# Fungsi sanitasi teks agar aman dimasukkan ke dokumen Word
def sanitize_text(text):
    # Hapus NULL byte dan karakter kontrol (selain newline dan tab)
    text = text.replace('\x00', '')
    text = re.sub(r'[\x01-\x08\x0b-\x1f\x7f]', '', text)
    return text.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="PDF / Gambar ke Word", layout="centered")
st.title("📄 PDF / Gambar ke Word Converter")
st.write("Upload file PDF atau gambar (JPG/PNG), lalu konversi ke file Word (.docx) secara otomatis.")

# Upload file
uploaded_file = st.file_uploader("Unggah file PDF atau Gambar", type=["pdf", "jpg", "jpeg", "png"])
use_ocr = st.checkbox("Gunakan OCR (untuk file hasil scan/gambar)")

# Tombol konversi
if uploaded_file and st.button("🔁 Convert to Word"):
    doc = Document()
    file_name = uploaded_file.name.lower()

    try:
        # Proses Gambar
        if file_name.endswith((".jpg", ".jpeg", ".png")):
            st.info("📷 Menggunakan OCR untuk gambar...")
            image = Image.open(uploaded_file)
            text = pytesseract.image_to_string(image)
            clean_text = sanitize_text(text)
            doc.add_paragraph(clean_text)
            st.success("✅ Konversi gambar selesai!")

        # Proses PDF
        elif file_name.endswith(".pdf"):
            if use_ocr:
                st.info("📄 PDF akan diproses dengan OCR...")
                images = convert_from_bytes(uploaded_file.read())
                for i, img in enumerate(images):
                    text = pytesseract.image_to_string(img)
                    clean_text = sanitize_text(text)
                    doc.add_paragraph(clean_text)
                    st.success(f"OCR halaman {i+1} selesai")
            else:
                st.info("📄 Mengekstrak teks langsung dari PDF...")
                uploaded_file.seek(0)
                with pdfplumber.open(uploaded_file) as pdf:
                    for i, page in enumerate(pdf.pages):
                        text = page.extract_text()
                        if text:
                            clean_text = sanitize_text(text)
                            doc.add_paragraph(clean_text)
                        st.success(f"Ekstraksi halaman {i+1} selesai")

        else:
            st.warning("❌ Jenis file tidak didukung.")
            st.stop()

        # Simpan ke .docx
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        file_base = uploaded_file.name.rsplit(".", 1)[0]
        st.download_button(
            label="📥 Download Word File",
            data=buffer,
            file_name=f"{file_base} (convert).docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"🚫 Terjadi kesalahan saat konversi: {e}")
