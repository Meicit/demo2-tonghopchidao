import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io

# --- 1. CẤU HÌNH ---
st.set_page_config(page_title="Hệ thống Quản lý Chỉ đạo", layout="wide")
st.title("🏛️ Trợ lý AI Hành chính (Bản v7 - Dứt điểm lỗi)")

if 'data' not in st.session_state:
    st.session_state['data'] = None

# --- 2. BỘ DÒ TÌM MODEL KHẢ DỤNG ---
def setup_ai():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Chưa cấu hình API Key trong Secrets!")
        return None
    
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    
    try:
        # Lấy danh sách model thực tế mà Key này được dùng
        valid_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        
        if not valid_models:
            st.error("❌ Key này không có quyền chạy bất kỳ Model nào!")
            return None
            
        # Ưu tiên lấy bản Flash cho nhanh
        target = next((m for m in valid_models if "1.5-flash" in m), valid_models[0])
        return genai.GenerativeModel(target)
    except Exception as e:
        st.error(f"❌ Lỗi kết nối API: {e}")
        return None

model = setup_ai()

# --- 3. XỬ LÝ DỮ LIỆU ---
def extract_full_text(file):
    try:
        if file.type == "application/pdf":
            return "\n".join([p.extract_text() for p in PdfReader(file).pages])
        return "\n".join([p.text for p in Document(file).paragraphs])
    except: return ""

def parse_table(text):
    try:
        lines = [l.strip() for l in text.split('\n') if '|' in l]
        data = [l for l in lines if not any(c in l for c in [':-', '---'])]
        if len(data) < 2: return pd.DataFrame([{"Kết quả": text}])
        cols = [c.strip() for c in data[0].split('|') if c.strip()]
        rows = [[c.strip() for c in l.split('|') if c.strip()] for l in data[1:]]
        return pd.DataFrame([r for r in rows if len(r) == len(cols)], columns=cols)
    except: return pd.DataFrame([{"Nội dung": text}])

# --- 4. GIAO DIỆN ---
file = st.file_uploader("Tải file PDF/Word", type=["pdf", "docx"])

if file and model:
    if st.button("🚀 TRÍCH XUẤT ĐẦY ĐỦ"):
        raw_text = extract_full_text(file)
        with st.spinner("AI đang quét văn bản..."):
            # Prompt cực ngắn để tránh lỗi quá tải
            prompt = f"Trích xuất tất cả nhiệm vụ thành bảng (STT | Nhiệm vụ | Đơn vị | Thời hạn). Liệt kê chi tiết, không tóm tắt:\n\n{raw_text[:10000]}"
            try:
                resp = model.generate_content(prompt)
                st.session_state['data'] = resp.text
            except Exception as e:
                st.error(f"Lỗi thực thi: {e}")

    if st.session_state['data']:
        st.markdown(st.session_state['data'])
        df = parse_table(st.session_state['data'])
        
        # Tạo file Excel chuẩn
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as w:
            df.to_excel(w, index=False)
            
        st.download_button("📥 Tải Excel đầy đủ", buf.getvalue(), f"Chi_dao_{file.name}.xlsx")
