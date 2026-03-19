import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io
import re

# --- CẤU HÌNH ---
st.set_page_config(page_title="AI Hành Chính v4", layout="wide")
st.title("🏛️ Hệ thống Trích xuất Toàn diện Chỉ đạo")

if 'analysis_result' not in st.session_state:
    st.session_state['analysis_result'] = None

# --- KẾT NỐI AI ---
def get_working_model():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Thiếu API Key!")
        return None
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    try:
        # Thử trực tiếp tên model flash để ổn định nhất
        return genai.GenerativeModel('gemini-1.5-flash')
    except: return None

model = get_working_model()

# --- HÀM CHUYỂN MARKDOWN THÀNH DATAFRAME (CẢI TIẾN) ---
def markdown_to_df(md_text):
    try:
        # Lọc ra các dòng có chứa dấu gạch đứng của bảng
        lines = [line.strip() for line in md_text.split('\n') if '|' in line]
        if len(lines) < 2: return pd.DataFrame([{"Nội dung": md_text}])
        
        # Loại bỏ dòng ngăn cách tiêu đề (|---|---|)
        data_lines = [l for l in lines if not re.match(r'^[|:\-\s]+$', l)]
        
        # Lấy tiêu đề và dữ liệu
        raw_columns = [c.strip() for c in data_lines[0].split('|') if c.strip()]
        rows = []
        for l in data_lines[1:]:
            row = [c.strip() for c in l.split('|') if c.strip()]
            if len(row) >= len(raw_columns):
                rows.append(row[:len(raw_columns)]) # Đảm bảo đúng số cột
        
        return pd.DataFrame(rows, columns=raw_columns)
    except Exception as e:
        return pd.DataFrame([{"Lỗi cấu trúc bảng": str(e), "Nội dung gốc": md_text}])

# --- GIAO DIỆN ---
uploaded_file = st.file_uploader("Tải lên file văn bản (PDF, DOCX)", type=["pdf", "docx"])

if uploaded_file and model:
    if st.button("🚀 BẮT ĐẦU TRÍCH XUẤT TOÀN BỘ"):
        # Đọc text
        text_content = ""
        if uploaded_file.type == "application/pdf":
            text_content = "\n".join([p.extract_text() for p in PdfReader(uploaded_file).pages])
        else:
            text_content = "\n".join([p.text for p in Document(uploaded_file).paragraphs])
            
        with st.spinner("AI đang quét toàn bộ văn bản..."):
            # Prompt yêu cầu KHÔNG ĐƯỢC TÓM TẮT, PHẢI LIÊT KÊ HẾT
            prompt = f"""
            Bạn là thư ký tổng hợp. Hãy đọc kỹ văn bản và trích xuất TẤT CẢ các chỉ đạo/nhiệm vụ.
            KHÔNG ĐƯỢC TÓM TẮT. PHẢI LIỆT KÊ ĐẦY ĐỦ các nhiệm vụ xuất hiện trong văn bản.
            Trình bày DUY NHẤT dưới dạng bảng Markdown: STT | Nhiệm vụ | Đơn vị thực hiện | Thời hạn.
            
            Văn bản:
            {text_content[:15000]}
            """
            response = model.generate_content(prompt)
            st.session_state['analysis_result'] = response.text

    if st.session_state['analysis_result']:
        st.markdown("### 📋 Danh sách chỉ đạo chi tiết")
        st.markdown(st.session_state['analysis_result'])
        
        # Chuyển đổi sang bảng DataFrame
        df = markdown_to_df(st.session_state['analysis_result'])
        
        st.subheader("📥 Xuất dữ liệu")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='NhiemVuChiTiet')
            
        st.download_button(
            label="Tải file Excel đầy đủ (.xlsx)",
            data=output.getvalue(),
            file_name=f"Chi_dao_chi_tiet_{uploaded_file.name}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
