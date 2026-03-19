import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io
import re

# --- CẤU HÌNH ---
st.set_page_config(page_title="AI Hành Chính v3", layout="wide")
st.title("🏛️ Hệ thống Trích xuất & Cấu trúc hóa Chỉ đạo")

if 'analysis_result' not in st.session_state:
    st.session_state['analysis_result'] = None

# --- KẾT NỐI AI ---
def get_working_model():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Thiếu API Key!")
        return None
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        priority = ['gemini-1.5-flash', 'models/gemini-1.5-flash']
        for p in priority:
            if p in available_models or p.replace('models/', '') in str(available_models):
                return genai.GenerativeModel(p)
        return genai.GenerativeModel(available_models[0])
    except: return None

model = get_working_model()

# --- HÀM CHUYỂN MARKDOWN THÀNH DATAFRAME ---
def markdown_to_df(md_text):
    try:
        # Tìm bảng trong chuỗi Markdown
        lines = [line.strip() for line in md_text.split('\n') if '|' in line]
        if len(lines) < 2: return pd.DataFrame([{"Nội dung": md_text}])
        
        # Loại bỏ dòng gạch ngang phân cách (|---|---|)
        data_lines = [l for l in lines if not re.match(r'^[|:\-\s]+$', l)]
        
        columns = [c.strip() for c in data_lines[0].split('|') if c.strip()]
        rows = []
        for l in data_lines[1:]:
            row = [c.strip() for c in l.split('|') if c.strip()]
            if len(row) == len(columns):
                rows.append(row)
        
        return pd.DataFrame(rows, columns=columns)
    except:
        return pd.DataFrame([{"Nội dung": md_text}])

# --- GIAO DIỆN ---
uploaded_file = st.file_uploader("Tải lên file văn bản", type=["pdf", "docx"])

if uploaded_file and model:
    if st.button("🚀 Phân tích & Cấu trúc dữ liệu"):
        # Đọc text (giữ nguyên hàm extract cũ của bạn)
        if uploaded_file.type == "application/pdf":
            content = "\n".join([p.extract_text() for p in PdfReader(uploaded_file).pages])
        else:
            content = "\n".join([p.text for p in Document(uploaded_file).paragraphs])
            
        with st.spinner("Đang bóc tách dữ liệu..."):
            prompt = f"Trích xuất nhiệm vụ thành bảng Markdown (STT | Nhiệm vụ | Đơn vị | Thời hạn). Lưu ý: Chỉ trả về bảng, không nói thêm.\n\n{content[:10000]}"
            response = model.generate_content(prompt)
            st.session_state['analysis_result'] = response.text

    if st.session_state['analysis_result']:
        st.markdown(st.session_state['analysis_result'])
        
        # Chuyển đổi sang bảng Excel xịn
        df = markdown_to_df(st.session_state['analysis_result'])
        
        st.subheader("📥 Tải file Excel chuẩn cột")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='DanhSachChiDao')
            
        st.download_button(
            label="Tải file Excel (.xlsx)",
            data=output.getvalue(),
            file_name=f"Danh_sach_nhiem_vu.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
