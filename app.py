import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io
import re

# --- CẤU HÌNH ---
st.set_page_config(page_title="AI Hành Chính Toàn Diện", layout="wide")
st.title("🏛️ Hệ thống Trích xuất Chỉ đạo (Bản đầy đủ)")

# Lưu kết quả vào bộ nhớ tạm
if 'full_analysis' not in st.session_state:
    st.session_state['full_analysis'] = None

# --- KẾT NỐI AI (SỬA LỖI NOTFOUND) ---
def get_model():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Thiếu API Key trong Secrets!")
        st.stop()
    
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    
    # Thử gọi model với tên chính xác nhất để tránh lỗi NotFound
    try:
        # Sử dụng 'gemini-1.5-flash' là bản ổn định và nhanh nhất hiện nay
        return genai.GenerativeModel('models/gemini-1.5-flash')
    except Exception as e:
        st.error(f"Lỗi khởi tạo Model: {e}")
        st.stop()

model = get_model()

# --- HÀM XỬ LÝ DỮ LIỆU ---
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
file = st.file_uploader("Tải lên văn bản (PDF, DOCX)", type=["pdf", "docx"])

if file:
    # Đọc văn bản
    if file.type == "application/pdf":
        raw_text = "\n".join([p.extract_text() for p in PdfReader(file).pages])
    else:
        raw_text = "\n".join([p.text for p in Document(file).paragraphs])

    if st.button("🚀 TRÍCH XUẤT TẤT CẢ NHIỆM VỤ"):
        with st.spinner("AI đang quét từng dòng văn bản..."):
            # Prompt được thiết kế để AI không dám tóm tắt
            prompt = f"""
            NHIỆM VỤ: Bạn là một thư ký tổng hợp có nhiệm vụ trích xuất KHÔNG SÓT MỘT CHI TIẾT NÀO.
            YÊU CẦU: 
            1. Liệt kê TẤT CẢ các nhiệm vụ, chỉ đạo, lời dặn của lãnh đạo trong văn bản.
            2. Không được tóm tắt chung chung. Một văn bản có 10 chỉ đạo phải liệt kê đủ 10 dòng.
            3. Trình bày dạng bảng Markdown: STT | Nội dung nhiệm vụ | Đơn vị thực hiện | Thời hạn.
            
            VĂN BẢN ĐẦU VÀO:
            {raw_text}
            """
            try:
                response = model.generate_content(prompt)
                st.session_state['full_analysis'] = response.text
            except Exception as e:
                st.error(f"Lỗi AI: {e}")

    # Hiển thị và Xuất dữ liệu
    if st.session_state['full_analysis']:
        st.markdown(st.session_state['full_analysis'])
        df = markdown_to_df(st.session_state['full_analysis'])
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
            
        st.download_button(
            label="📥 Tải về file Excel đầy đủ",
            data=output.getvalue(),
            file_name=f"Ket_qua_chi_tiet.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
