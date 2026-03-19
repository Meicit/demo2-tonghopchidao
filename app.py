import streamlit as st
import google.generativeai as genai
from pypdf import PdfReader
from docx import Document
import pandas as pd
import io

# 1. Cấu hình trang
st.set_page_config(page_title="Hệ thống AI Hành chính", layout="wide")
st.title("🏛️ Trợ lý Trích xuất & Xuất Excel")

# 2. Khởi tạo AI
if "GEMINI_API_KEY" not in st.secrets:
    st.error("Vui lòng thêm GEMINI_API_KEY vào mục Secrets!")
    st.stop()

genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
model = genai.GenerativeModel('gemini-1.5-flash')

# 3. Xử lý file
file = st.file_uploader("Tải lên PDF hoặc Word", type=["pdf", "docx"])

if file:
    text = ""
    try:
        if file.type == "application/pdf":
            reader = PdfReader(file)
            text = "\n".join([p.extract_text() for p in reader.pages if p.extract_text()])
        else:
            doc = Document(file)
            text = "\n".join([p.text for p in doc.paragraphs])
    except Exception as e:
        st.error(f"Lỗi đọc file: {e}")

    if st.button("🚀 Phân tích & Trích xuất"):
        if not text:
            st.warning("Không thể đọc nội dung file.")
        else:
            with st.spinner("AI đang xử lý..."):
                try:
                    # Prompt yêu cầu bảng đơn giản
                    prompt = f"Trích xuất tất cả nhiệm vụ thành bảng (STT | Nhiệm vụ | Đơn vị | Thời hạn). Liệt kê đầy đủ, không tóm tắt:\n\n{text[:15000]}"
                    response = model.generate_content(prompt)
                    res_text = response.text
                    
                    st.markdown("### ✅ Kết quả phân tích")
                    st.markdown(res_text)
                    
                    # Tạo file Excel từ kết quả
                    output = io.BytesIO()
                    # Chia nhỏ text thành dòng để Excel dễ nhìn hơn
                    df_export = pd.DataFrame([{"Nội dung trích xuất": res_text}])
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_export.to_excel(writer, index=False)
                    
                    st.download_button(
                        label="📥 Tải về Excel",
                        data=output.getvalue(),
                        file_name=f"Trich_xuat_{file.name}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"Lỗi AI: {e}")
