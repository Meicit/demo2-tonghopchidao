import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io

# --- 1. CẤU HÌNH GIAO DIỆN ---
st.set_page_config(page_title="AI Hành Chính", layout="wide")
st.title("🏛️ Hệ thống Trích xuất & Xuất Excel")

# --- 2. KẾT NỐI AI (CƠ CHẾ TỰ QUÉT MODEL) ---
def get_working_model():
    if "GEMINI_API_KEY" not in st.secrets:
        st.error("❌ Thiếu API Key! Hãy thêm vào Secrets.")
        return None
    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
    try:
        available_models = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        # Thử gọi tên model không có prefix 'models/' trước vì đôi khi nó gây lỗi 404
        priority = ['gemini-1.5-flash', 'models/gemini-1.5-flash', 'gemini-1.0-pro']
        for p in priority:
            if p in available_models or p.replace('models/', '') in str(available_models):
                return genai.GenerativeModel(p)
        return genai.GenerativeModel(available_models[0]) if available_models else None
    except Exception as e:
        st.error(f"❌ Lỗi kết nối API: {e}")
        return None

model = get_working_model()

# --- 3. HÀM ĐỌC FILE ---
def extract_text(file):
    text = ""
    try:
        if file.type == "application/pdf":
            reader = PdfReader(file)
            for page in reader.pages:
                text += page.extract_text() + "\n"
        else:
            doc = Document(file)
            text = "\n".join([p.text for p in doc.paragraphs])
        return text.strip()
    except: return ""

# --- 4. GIAO DIỆN NGƯỜI DÙNG ---
uploaded_file = st.file_uploader("Tải lên file PDF hoặc Word", type=["pdf", "docx"])

if uploaded_file and model:
    content = extract_text(uploaded_file)
    if st.button("🚀 Bắt đầu Phân tích & Tạo file tải về"):
        with st.spinner("AI đang xử lý..."):
            # Prompt yêu cầu bảng Markdown truyền thống (Rất ổn định)
            prompt = f"""
            Bạn là trợ lý hành chính. Hãy đọc văn bản sau và trích xuất thành bảng Markdown gồm: 
            STT, Nội dung nhiệm vụ, Đơn vị thực hiện, Thời hạn. 
            Văn bản: {content[:10000]}
            """
            try:
                response = model.generate_content(prompt)
                res_text = response.text
                st.markdown("### ✅ Kết quả trích xuất:")
                st.markdown(res_text)

                # --- XỬ LÝ XUẤT FILE (BẢN AN TOÀN) ---
                # Lưu kết quả vào file Excel đơn giản
                output = io.BytesIO()
                df_export = pd.DataFrame([{"Nội dung chỉ đạo": res_text}])
                
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_export.to_excel(writer, index=False, sheet_name='KetQua')
                
                st.divider()
                st.subheader("📥 Tải dữ liệu về máy")
                col1, col2 = st.columns(2)
                
                with col1:
                    st.download_button(
                        label="Excel (.xlsx)",
                        data=output.getvalue(),
                        file_name=f"Chi_dao_{uploaded_file.name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                with col2:
                    st.download_button(
                        label="Văn bản (.txt)",
                        data=res_text,
                        file_name=f"Chi_dao_{uploaded_file.name}.txt",
                        mime="text/plain"
                    )
            except Exception as e:
                st.error(f"Lỗi thực thi: {e}")
