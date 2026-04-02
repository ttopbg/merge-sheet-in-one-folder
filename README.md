# 📚 Gộp danh sách học sinh – Excel Merger

Ứng dụng Streamlit tự động nhận diện và gộp nhiều file Excel danh sách học sinh thành một file duy nhất.

## Tính năng
- Upload nhiều file `.xlsx`, `.xls`, `.xlsm` cùng lúc
- Tự nhận diện cột **Lớp**, **Họ và tên**, **Ngày sinh** dù tên cột khác nhau (Lớp học, Class, Họ tên, Full name, DOB, Ngày tháng năm sinh…)
- Đọc tất cả các sheet trong mỗi file
- Gộp nối tiếp theo cột, loại trùng theo (Họ tên + Ngày sinh)
- Tên file output mặc định = tên file đầu vào + `_đã gộp`

---

## Deploy lên Streamlit Cloud (miễn phí)

### Bước 1 – Tạo repo GitHub
1. Vào [github.com](https://github.com) → **New repository**
2. Đặt tên repo (ví dụ: `gop-danh-sach`)
3. Chọn **Public** → **Create repository**

### Bước 2 – Upload các file sau vào repo
```
app.py
merge_excel.py
requirements.txt
```
(Kéo thả trực tiếp vào trang GitHub của repo)

### Bước 3 – Deploy trên Streamlit Cloud
1. Vào [share.streamlit.io](https://share.streamlit.io) → đăng nhập bằng GitHub
2. Nhấn **New app**
3. Chọn repo `gop-danh-sach`, branch `main`, file `app.py`
4. Nhấn **Deploy!**

Sau vài phút bạn có link kiểu:  
`https://gop-danh-sach.streamlit.app`

---

## Chạy thử trên máy local

```bash
pip install -r requirements.txt
streamlit run app.py
```
