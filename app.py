import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io
import re

# --- CẤU HÌNH ---
st.set_page_config(page_title="AI Hành Chính v6", layout="wide")
st.title("🏛️ Hệ thống Trích xuất Chỉ đạo (Bản Fix 404)")

# Lưu kết quả
if 'final_data' not in st.session_state:
    st.session_state['final_data'] = None

# --- KẾT NỐI AI (CÁCH GỌI TRỰC TIẾP NHẤT) ---
def init_genai():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Thiếu API Key trong Secrets!")
        st.stop()
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])

init_genai()

# --- HÀM XỬ LÝ ---
def markdown_to_df(md_text):
    try:
        lines = [l.strip() for l in md_text.split('\n') if '|' in l]
        if len(lines) < 2: return pd.DataFrame([{"Nội dung": md_text}])
        data_lines = [l for l in lines if not re.match(r'^[|:\-\s]+$', l)]
        cols = [c.strip() for c in data_lines[0].split('|') if c.strip()]
        rows = []
        for l in data_lines[1:]:
            row = [c.strip() for c in l.split('|') if c.strip()]
            if len(row) >= len(cols): rows.append(row[:len(cols)])
        return pd.DataFrame(rows, columns=cols)
    except: return pd.DataFrame([{"Nội dung": md_text}])

# --- GIAO DIỆN ---
file = st.file_uploader("Tải lên file văn bản", type=["pdf", "docx"])

if file:
    if file.type == "application/pdf":
        raw_text = "\n".join([p.extract_text() for p in PdfReader(file).pages])
    else:
        raw_text = "\n".join([p.text for p in Document(file).paragraphs])

    if st.button("🚀 TRÍCH XUẤT DỮ LIỆU"):
        with st.spinner("AI đang xử lý..."):
            # THAY ĐỔI QUAN TRỌNG: Gọi model trực tiếp trong hàm để tránh mất kết nối
            try:
                # Thử dùng tên ngắn gọn nhất - thường là bản ổn định nhất
                model = genai.GenerativeModel('gemini-1.5-flash') 
                
                prompt = f"Trích xuất tất cả nhiệm vụ thành bảng Markdown (STT | Nhiệm vụ | Đơn vị | Thời hạn):\n\n{raw_text[:12000]}"
                response = model.generate_content(prompt)
                
                if response:
                    st.session_state['final_data'] = response.text
            except Exception as e:
                # Nếu vẫn lỗi, thử dùng tên có prefix 'models/'
                try:
                    model = genai.GenerativeModel('models/gemini-1.5-flash')
                    response = model.generate_content(prompt)
                    st.session_state['final_data'] = response.text
                except Exception as e2:
                    st.error(f"Lỗi hệ thống Google AI: {e2}")

    if st.session_state['final_data']:
        st.markdown(st.session_state['final_data'])
        df = markdown_to_df(st.session_state['final_data'])
        
        # Xuất Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button("📥 Tải về Excel", output.getvalue(), "chidao.xlsx")
