import streamlit as st
import pandas as pd
from merge_excel import merge_excel_files, to_excel_bytes

st.set_page_config(
    page_title="Gộp HS trong Folder",
    page_icon="🐍",
    layout="centered",
)

# ─── CSS tuỳ chỉnh ───────────────────────────────────────────────────────────
st.markdown("""
<style>
    .main { max-width: 820px; margin: auto; }
    .stAlert { border-radius: 8px; }
    h1 { color: #1a5276; }
    .log-box {
        background: #f4f6f7;
        border-left: 4px solid #2e86c1;
        padding: 10px 16px;
        border-radius: 6px;
        font-size: 0.88rem;
        font-family: monospace;
        max-height: 260px;
        overflow-y: auto;
    }
</style>
""", unsafe_allow_html=True)

# ─── Tiêu đề ─────────────────────────────────────────────────────────────────
st.title("🎲 Gộp học sinh từ nhiều file trong một Folder")
st.markdown(
    "Upload nhiều file Excel (`.xlsx`, `.xls`, `.xlsm`). "
    "Ứng dụng tự nhận diện cột **Lớp**, **Họ và tên**, **Ngày sinh** "
    "rồi gộp tất cả thành một file duy nhất."
)

# ─── Upload ──────────────────────────────────────────────────────────────────
uploaded_files = st.file_uploader(
    "Chọn một hoặc nhiều file Excel",
    type=["xlsx", "xls", "xlsm"],
    accept_multiple_files=True,
)

output_name_placeholder = "TongHop_đã_gộp"
if uploaded_files:
    if len(uploaded_files) == 1:
        base = uploaded_files[0].name.rsplit(".", 1)[0]
        output_name_placeholder = f"{base}_đã gộp"
    else:
        output_name_placeholder = "DanhSach_đã gộp"

output_name = st.text_input(
    "Tên file kết quả (không cần đuôi .xlsx)",
    value=output_name_placeholder,
)

# ─── Nút xử lý ───────────────────────────────────────────────────────────────
if uploaded_files:
    if st.button("▶️  Gộp file", use_container_width=True, type="primary"):
        with st.spinner("Đang xử lý..."):
            merged_df, logs = merge_excel_files(uploaded_files)

        # Log
        st.markdown("**Nhật ký xử lý:**")
        log_html = "<div class='log-box'>" + "<br>".join(logs) + "</div>"
        st.markdown(log_html, unsafe_allow_html=True)

        if merged_df.empty:
            st.error("Không tìm thấy dữ liệu hợp lệ trong các file đã upload.")
        else:
            st.success(f"✅ Tổng cộng **{len(merged_df)}** học sinh sau khi gộp và loại trùng.")

            # Xem trước
            st.markdown("**Xem trước dữ liệu:**")
            st.dataframe(merged_df.head(50), use_container_width=True)

            # Tải xuống
            excel_bytes = to_excel_bytes(merged_df)
            safe_name = output_name.strip() or "TongHop"
            # Bỏ đuôi nếu user đã nhập
            if safe_name.lower().endswith(".xlsx"):
                safe_name = safe_name[:-5]
            download_filename = f"{safe_name}.xlsx"

            st.download_button(
                label="⬇️  Tải file Excel tổng hợp",
                data=excel_bytes,
                file_name=download_filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
else:
    st.info("⬆️ Hãy upload ít nhất một file Excel để bắt đầu.")

# ─── Footer ──────────────────────────────────────────────────────────────────
st.divider()
st.caption(
    "Cột được nhận diện tự động: **Lớp** (Lớp học, Class…) · "
    "**Họ và tên** (Tên, Full name…) · **Ngày sinh** (Ngày tháng năm sinh, DOB…). "
    "Dữ liệu trùng (cùng họ tên + ngày sinh) sẽ được loại bỏ."
)
