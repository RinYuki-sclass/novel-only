# AI Novel Translation Tool (Trans-Tool) Web

Công cụ hỗ trợ dịch thuật tiểu thuyết tự động sử dụng sức mạnh của **Google Gemini AI**, được thiết kế thông minh với giao diện Web (Streamlit) để hoạt động hiệu quả cho làm việc nhóm. 

Tối ưu cho việc dịch truyện chữ (từ tiếng Anh/Hàn) và truyện tranh (Manhwa) sang tiếng Việt, tích hợp hệ thống quản lý Thuật ngữ (Glossary) và Kiểm soát chất lượng (QC) nghiêm ngặt.

---

## 🚀 Tính năng nổi bật
- **Dịch thuật 2 bước (Translate & Refine):** Nhận song song English & Korean để dịch và tinh chỉnh văn phong, cực kỳ chuẩn xác.
- **Dịch Truyện Tranh (Manhwa OCR):** Quét bong bóng thoại từ hình ảnh, tự động dịch và trả về văn bản sắp xếp theo luồng đọc.
- **Hệ thống QC Review:** AI tự check lỗi dịch thuật, sai Glossary, hoặc thiếu sót so với bản gốc. Tự động gợi ý Thuật ngữ mới.
- **Trình Đối Chiếu (Side-by-Side):** Giao diện tương tác cho phép xem bản dịch song song với bản gốc và chỉnh sửa/lưu trực tiếp.
- **Auto Model & Fallback (Thông minh):** Tự động điều phối các Model AI chuyên biệt (`gemini-3-flash-preview` cho Dịch, `gemini-2.5-flash` cho QC) và tự động Fallback về `gemini-3.1-flash-lite` (500 RPD) khi hết hạn mức để không cản trở công việc.
- **Quản lý đa API Keys (Anti-Rate Limit):** Hỗ trợ khai báo tới 20 API Keys. Tự động chuyển Key khi một Key chạm mức RPD (Requests Per Day).

---

## 🛠 Cài đặt & Khởi chạy (Dành cho máy Local)

### 1. Yêu cầu hệ thống
- Cài đặt [Python 3.9+](https://www.python.org/downloads/).
- Khởi chạy Terminal/CMD tại thư mục dự án và cài đặt thư viện:
  ```bash
  pip install -r requirements.txt
  ```

### 2. Cấu hình file `.env`
Tạo file `.env` tại thư mục gốc (ngang hàng `README.md`) với format sau:
```env
# Mật khẩu API của bạn
GEMINI_API_KEY_1=your_api_key_1_here
GEMINI_API_KEY_2=your_api_key_2_here

# Ẩn nút chọn "Dữ liệu từ File" (chỉ cho phép Copy Paste) trên Web
HIDE_LOCAL_FILE_OPTION=False
```

### 3. Chạy Website
Gõ lệnh này vào Terminal để mở Web App:
```bash
streamlit run scripts/app.py
```

---

## 🌐 Hướng dẫn Host Website lên Streamlit Cloud (Miễn phí cho Team)

1. Push toàn bộ mã nguồn này lên một kho lưu trữ Github (Private/Public đều được).
   > **Lưu ý:** KHÔNG push file `.env`. File này đã được đưa vào `.gitignore`.
2. Truy cập [Streamlit Community Cloud](https://share.streamlit.io/), đăng nhập và liên kết với Github.
3. Tạo New App -> Chọn repository của bạn -> Đặt `Main file path` là `scripts/app.py`.
4. Nhấn **Deploy**.
5. Trong cấu hình App trên Streamlit Dashboard, chọn **Settings** > **Secrets**. Dán cấu hình hệ thống bằng định dạng TOML:
   ```toml
   GEMINI_API_KEY_1="your_api_key_1_here"
   GEMINI_API_KEY_2="your_api_key_2_here"

   # Khi Host public cho Team, bật True để ép mọi người dùng chế độ Copy-Paste
   HIDE_LOCAL_FILE_OPTION="True"
   ```
6. Bấm Save. Vậy là team bạn đã có một công cụ xịn sò, dùng chung dung lượng Rate Limit hiển thị mượt mà trên UI.

---

## 📂 Cách cấu hình Từ Điển (Glossary)
Từ điển (Glossary) dùng chung cực kỳ quan trọng, là "não bộ" của con AI.
* `glossary.md`: Lưu định nghĩa Nhân vật, Phái, Tuyệt chiêu, Địa danh... Định dạng bảng hướng dẫn chi tiết nằm bên trong.
* `personal_notes.md`: Nơi Lead quy định cách xưng hô đặc biệt hoặc văn phong bắt buộc (VD: Không được chêm quá nhiều từ Hán Việt, Xưng hô Anh - Em...).