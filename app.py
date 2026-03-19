import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io

# --- 1. CẤU HÌNH ---
st.set_page_config(page_title="AI Hành Chính v2", layout="wide")
st.title("🏛️ Hệ thống Trích xuất Chỉ đạo Văn bản")

# Khởi tạo Session State để lưu kết quả
if 'analysis_result' not in st.session_state:
    st.session_state['analysis_result'] = None

# --- 2. KẾT NỐI AI ---
def get_working_model():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Thiếu API Key trong Secrets!")
        return None
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    try:
        # Cơ chế quét model an toàn
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        priority = ['gemini-1.5-flash', 'models/gemini-1.5-flash', 'gemini-1.0-pro']
        for p in priority:
            if p in available_models or p.replace('models/', '') in str(available_models):
                return genai.GenerativeModel(p)
        return genai.GenerativeModel(available_models[0]) if available_models else None
    except Exception as e:
        st.error(f"❌ Lỗi kết nối API: {e}")
        return None

model = get_working_model()

# --- 3. XỬ LÝ FILE ---
def extract_text(file):
    try:
        if file.type == "application/pdf":
            reader = PdfReader(file)
            return "\n".join([page.extract_text() for page in reader.pages])
        else:
            doc = Document(file)
            return "\n".join([p.text for p in doc.paragraphs])
    except: return ""

# --- 4. GIAO DIỆN ---
uploaded_file = st.file_uploader("Tải lên file văn bản", type=["pdf", "docx"])

if uploaded_file and model:
    # Nếu người dùng tải file mới, xóa kết quả cũ
    if 'last_uploaded' not in st.session_state or st.session_state['last_uploaded'] != uploaded_file.name:
        st.session_state['analysis_result'] = None
        st.session_state['last_uploaded'] = uploaded_file.name

    if st.button("🚀 Bắt đầu Phân tích"):
        content = extract_text(uploaded_file)
        with st.spinner("AI đang xử lý..."):
            prompt = f"Trích xuất nhiệm vụ từ văn bản sau thành bảng Markdown (STT, Nhiệm vụ, Đơn vị, Thời hạn):\n\n{content[:10000]}"
            try:
                response = model.generate_content(prompt)
                # LƯU KẾT QUẢ VÀO SESSION STATE
                st.session_state['analysis_result'] = response.text
            except Exception as e:
                st.error(f"Lỗi: {e}")

    # HIỂN THỊ KẾT QUẢ (Nếu đã có trong bộ nhớ tạm)
    if st.session_state['analysis_result']:
        st.divider()
        st.markdown("### ✅ Kết quả trích xuất:")
        st.markdown(st.session_state['analysis_result'])

        # --- XỬ LÝ TẢI VỀ ---
        # Tạo file Excel từ kết quả lưu trong session_state
        output = io.BytesIO()
        df_export = pd.DataFrame([{"Nội dung": st.session_state['analysis_result']}])
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df_export.to_excel(writer, index=False)
        
        st.subheader("📥 Tải dữ liệu về máy")
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="Tải file Excel (.xlsx)",
                data=output.getvalue(),
                file_name=f"Chi_dao_{uploaded_file.name}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        with col2:
            st.download_button(
                label="Tải file Văn bản (.txt)",
                data=st.session_state['analysis_result'],
                file_name=f"Chi_dao_{uploaded_file.name}.txt",
                mime="text/plain"
            )
