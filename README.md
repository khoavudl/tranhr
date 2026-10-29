# TranHR - Công Cụ Tự Động Đối Soát Thuế (MST & NPT)

Ứng dụng desktop viết bằng Python (giao diện Tkinter) giúp tự động hóa quá trình đối soát và cập nhật Mã số thuế từ các thông báo PDF của Cơ quan Thuế vào file Excel quản lý của Doanh nghiệp.

## 🌟 Tính Năng Chính

Ứng dụng được thiết kế theo dạng **Tab** gồm 2 công cụ độc lập:

1. **🏢 Tab 1: Cấp Mã Số Thuế (MST)**
   - Đối soát danh sách người lao động đăng ký MST cá nhân.
   - Đọc số CCCD tại **Cột M** (từ dòng 4) của file Excel (`Tong hop MST.xlsx`).
   - Cập nhật MST vào **Cột D**.
   - Bóc tách kết quả từ:
     - **Mẫu 14-MST**: Đánh dấu `CAP_MOI`, tô màu **Xanh lá**.
     - **Mẫu 40/TB-ĐKT**: Đánh dấu `TRUNG_MST`, bóc tách MST cũ 10 số, tô màu **Vàng**.
   - Tự động bổ sung 3 cột cuối bảng: `TRANG_THAI`, `GHI_CHU`, `FILE_NGUON`.

2. **👨‍👩‍👧 Tab 2: Người Phụ Thuộc (NPT)**
   - Đối soát danh sách đăng ký người phụ thuộc giảm trừ gia cảnh.
   - Đọc số CCCD tại **Cột K** (từ dòng 4) của file Excel (`Tong hop NPT.xlsx`).
   - Cập nhật MST vào **Cột D**.
   - **Tính năng độc quyền (Split-match)**: Tự động ghép 2 mảnh chuỗi số CCCD nếu bị ngắt trang (Page Break) trong file PDF của Cục Thuế.
   - Ưu tiên kiểm tra Mẫu 14 trước, sau đó tra sang Mẫu 40 nếu không có.

---

## 📁 Cấu Trúc Dự Án (Modular)

```text
TranHR/
├── core/
│   ├── __init__.py
│   ├── mst_processor.py       # Logic bóc tách & cập nhật MST
│   └── npt_processor.py       # Logic bóc tách & cập nhật NPT
├── ui/
│   ├── __init__.py
│   ├── app.py                 # Cửa sổ chính & thiết lập 2 Tabs
│   └── components.py          # Component giao diện, chọn file, console log
├── main.py                    # Điểm khởi chạy chính
├── run.sh                     # Script khởi động nhanh
├── requirements.txt           # Danh sách thư viện phụ thuộc
├── MST.ipynb                  # File Colab gốc (tham khảo)
└── NPT.ipynb                  # File Colab gốc (tham khảo)
```

---

## 🚀 Hướng Dẫn Sử Dụng

### Cách 1: Khởi động nhanh (Khuyên dùng)
Chỉ cần chạy lệnh sau trong Terminal:
```bash
./run.sh
```

### Cách 2: Khởi động thủ công qua Python
```bash
# Kích hoạt môi trường ảo
source venv/bin/activate

# Khởi chạy app
python3 main.py
```

---

## 💡 Hướng Dẫn Thao Tác Trên Giao Diện

1. **Chọn Tab nghiệp vụ**: Chọn tab `Cấp Mã Số Thuế (MST)` hoặc `Người Phụ Thuộc (NPT)`.
2. **Chọn file Excel**: Bấm nút `📂 Chọn file Excel...` để trỏ tới file Excel mẫu danh sách.
3. **Chọn file PDF**:
   - Bấm `📄 Chọn nhiều file PDF...` để chọn các file lẻ.
   - Hoặc bấm `📁 Chọn cả thư mục chứa PDF...` để phần mềm tự động nạp toàn bộ file PDF trong thư mục đó.
4. **Bấm `🚀 BẮT ĐẦU XỬ LÝ`**:
   - Phần mềm sẽ chạy ở luồng nền (không đơ máy), hiển thị tiến trình và log chi tiết.
   - Sau khi hoàn thành, ứng dụng sẽ hiện hộp thoại thông báo kết quả tổng kết (Cấp mới, Trùng MST, Chưa tìm thấy).
5. **Mở kết quả**:
   - Bấm `📊 Mở file Excel kết quả` để xem trực tiếp.
   - Bấm `📂 Mở thư mục kết quả` để chuyển đến vị trí lưu file.
