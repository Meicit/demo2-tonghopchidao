import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io

# --- 1. Cấu hình ---
st.set_page_config(page_title="AI Hành Chính", layout="wide")
st.title("🏛️ Hệ thống Trích xuất Chỉ đạo")

# Khởi tạo model an toàn
def get_model():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("Chưa có API Key!")
        return None
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    # Thử dùng tên model cơ bản nhất
    return genai.GenerativeModel('gemini-1.5-flash')

model = get_model()

# --- 2. Xử lý file ---
uploaded_file = st.file_uploader("Tải lên PDF/Word", type=["pdf", "docx"])

if uploaded_file and model:
    # Đọc văn bản
    text = ""
    if uploaded_file.type == "application/pdf":
        text = "\n".join([p.extract_text() for p in PdfReader(uploaded_file).pages])
    else:
        text = "\n".join([p.text for p in Document(uploaded_file).paragraphs])

    if st.button("🚀 Trích xuất & Tạo file Excel"):
        with st.spinner("AI đang xử lý..."):
            try:
                prompt = f"Trích xuất tất cả nhiệm vụ thành bảng (STT | Nhiệm vụ | Đơn vị | Thời hạn). Liệt kê chi tiết:\n\n{text[:10000]}"
                response = model.generate_content(prompt)
                res_text = response.text
                
                st.markdown(res_text)
                
                # Tạo file Excel đơn giản để tránh lỗi định dạng
                output = io.BytesIO()
                df_export = pd.DataFrame([{"Nội dung": res_text}])
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False)
                
                st.download_button("📥 Tải về Excel", output.getvalue(), "chidao.xlsx")
            except Exception as e:
                st.error(f"Lỗi: {e}")
