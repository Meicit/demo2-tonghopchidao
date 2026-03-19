import streamlit as st
from google import genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io

# --- 1. CẤU HÌNH ---
st.set_page_config(page_title="AI Hành Chính 2026", layout="wide")
st.title("🏛️ Hệ thống Trích xuất Chỉ đạo (Bản Cập nhật mới nhất)")

# Khởi tạo Client với thư viện mới
if "GEMINI_API_KEY" not in st.secrets:
    st.error("❌ Thiếu API Key trong Secrets!")
    st.stop()

client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])

# Lưu kết quả
if 'data' not in st.session_state:
    st.session_state['data'] = None

# --- 2. HÀM ĐỌC FILE ---
def extract_text(file):
    try:
        if file.type == "application/pdf":
            return "\n".join([p.extract_text() for p in PdfReader(file).pages if p.extract_text()])
        return "\n".join([p.text for p in Document(file).paragraphs])
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")
        return ""

# --- 3. GIAO DIỆN ---
uploaded_file = st.file_uploader("Tải lên văn bản (PDF, DOCX)", type=["pdf", "docx"])

if uploaded_file:
    if st.button("🚀 BẮT ĐẦU TRÍCH XUẤT"):
        text_content = extract_text(uploaded_file)
        if text_content:
            with st.spinner("AI đang xử lý bằng công nghệ mới nhất..."):
                try:
                    # Cách gọi mới của thư viện google-genai
                    response = client.models.generate_content(
                        model="gemini-1.5-flash",
                        contents=f"Trích xuất tất cả nhiệm vụ thành bảng (STT | Nhiệm vụ | Đơn vị | Thời hạn). Liệt kê đầy đủ, không tóm tắt:\n\n{text_content[:15000]}"
                    )
                    st.session_state['data'] = response.text
                except Exception as e:
                    st.error(f"Lỗi AI: {e}")

    if st.session_state['data']:
        st.markdown("### ✅ Kết quả trích xuất")
        st.markdown(st.session_state['data'])
        
        # Tạo file Excel đơn giản
        df_export = pd.DataFrame([{"Nội dung": st.session_state['data']}])
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 Tải về Excel",
            data=output.getvalue(),
            file_name=f"Chi_dao_{uploaded_file.name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
