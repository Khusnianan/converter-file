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
    text = text.replace('\x00', '')  # Hapus NULL byte
    text = re.sub(r'[\x01-\x08\x0b-\x1f\x7f]', '', text)  # Hapus karakter kontrol
    return text.encode('utf-8', errors='ignore').decode('utf-8', errors='ignore')

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="PDF / Gambar ke Word", layout="centered")

# Header cantik
st.markdown("""
<h1 style="text-align: center; color: #2E8B57;">📄 PDF / Gambar ➜ Word Converter</h1>
<p style="text-align: center; font-size: 18px;">Mudah mengubah file <strong>PDF, JPG, PNG</strong> menjadi <strong>Word (.docx)</strong>!</p>
<hr>
""", unsafe_allow_html=True)

# Upload
st.markdown("### 📁 Unggah File")
uploaded_file = st.file_uploader("Pilih file PDF, JPG, atau PNG untuk dikonversi.", type=["pdf", "jpg", "jpeg", "png"])

# OCR info
st.markdown("#### 🧠 Apakah ini file hasil scan atau gambar?")
use_ocr = st.checkbox("Aktifkan OCR (untuk file hasil scan/foto/gambar)", value=False)
st.caption("✅ Gunakan OCR jika file hasil scan atau tidak bisa disalin teksnya.")

st.markdown("---")

# Tombol konversi
if uploaded_file and st.button("🔁 Convert to Word"):
    doc = Document()
    file_name = uploaded_file.name.lower()

    try:
        # Gambar
        if file_name.endswith((".jpg", ".jpeg", ".png")):
            st.info("📷 Memproses gambar dengan OCR...")
            image = Image.open(uploaded_file)
            text = pytesseract.image_to_string(image)
            clean_text = sanitize_text(text)
            doc.add_paragraph(clean_text)
            st.success("✅ Konversi gambar selesai!")

        # PDF
        elif file_name.endswith(".pdf"):
            if use_ocr:
                st.info("📄 Memproses PDF dengan OCR...")
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

        # Simpan ke file .docx
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)

        # Nama file output
        file_base = uploaded_file.name.rsplit(".", 1)[0]
        output_filename = f"{file_base} (konversi).docx"

        # Download button
        st.download_button(
            label="📥 Download Word File",
            data=buffer,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"🚫 Terjadi kesalahan saat konversi: {e}")
