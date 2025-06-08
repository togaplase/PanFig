import streamlit as st
import pytesseract
import pandas as pd
import numpy as np
import cv2
from PIL import Image
import io

# Jika perlu, set path tesseract (lokal saja, tidak perlu di Streamlit Cloud)
# pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

st.set_page_config(page_title="Image Table to Excel", layout="centered")
st.title("📄 Gambar Tabel ➜ Excel")
st.markdown("Upload gambar tabel, lalu lihat & edit hasilnya sebelum mengunduh file Excel.")

uploaded_file = st.file_uploader("Unggah Gambar Tabel", type=["jpg", "jpeg", "png"])

if uploaded_file:
    st.image(uploaded_file, caption="Gambar diunggah", use_column_width=True)

    image = Image.open(uploaded_file).convert("RGB")
    img_array = np.array(image)
    gray = cv2.cvtColor(img_array, cv2.COLOR_RGB2GRAY)
    _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)

    with st.spinner("🔍 Mengekstrak teks dari gambar..."):
        data = pytesseract.image_to_string(thresh, config='--oem 3 --psm 6')

    rows = [row.strip() for row in data.split('\n') if row.strip()]
    table = [row.split('\t') for row in rows]
    df = pd.DataFrame(table)

    edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic")

    towrite = io.BytesIO()
    edited_df.to_excel(towrite, index=False, header=False, engine='openpyxl')
    towrite.seek(0)

    st.download_button(
        label="💾 Unduh sebagai Excel",
        data=towrite,
        file_name="tabel_ekstrak.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
