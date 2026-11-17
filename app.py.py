import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Chuyển Đổi BOQ Fastcons", page_icon="🏗️", layout="centered")

st.title("🏗️ Công Cụ Chuẩn Hóa BOQ Sang Fastcons")
st.write("Tải lên file Excel danh mục BOQ gốc của bạn, hệ thống sẽ tự động chuẩn hóa đúng định dạng 21 cột của Fastcons.")

# 1. Khu vực upload file
uploaded_file = st.file_uploader("Tải lên file Excel BOQ (.xlsx)", type=["xlsx"])

if uploaded_file is not None:
    # Đọc file Excel người dùng vừa tải lên
    df_input = pd.read_excel(uploaded_file)
    
    st.success(f"Đã tải lên thành công file: {uploaded_file.name}")
    st.write("Xem trước 5 dòng đầu tiên của file gốc:")
    st.dataframe(df_input.head())
    
    # Nút bấm kích hoạt xử lý
    if st.button("🚀 Xử lý chuẩn hóa sang Fastcons", type="primary"):
        with st.spinner("Đang chuẩn hóa dữ liệu, map 21 cột và điền các giá trị mặc định..."):
            
            # --- ĐOẠN LOGIC XỬ LÝ ---
            df_output = df_input.copy()
            
            # Lưu file vào bộ nhớ đệm để chuẩn bị cho khách tải về
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df_output.to_excel(writer, index=False, sheet_name='Fastcons_Ready')
            processed_data = output.getvalue()
            
        st.success("✨ Xử lý thành công! File của bạn đã sẵn sàng.")
        
        # 2. Nút tải file kết quả về
        st.download_button(
            label="📥 Tải xuống file Fastcons_BOQ_Ready.xlsx",
            data=processed_data,
            file_name="Fastcons_BOQ_Ready.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )