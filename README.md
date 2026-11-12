# 🚀 AutoOffice — Trợ lý AI thông minh cho Microsoft Word, Excel & PowerPoint

**AutoOffice** là tiện ích (Add-in) tích hợp trực tiếp vào thanh công cụ của **Microsoft Word, Excel và PowerPoint**, trang bị **Bộ 10 nhóm công cụ chuyên sâu (52 tác vụ)** giúp soạn thảo văn bản hành chính Việt Nam chuẩn Nghị định 30/2020/NĐ-CP, nhận diện trực tiếp vùng văn bản đang bôi chọn/in đậm/tô màu, và chạy đa luồng độc lập trên từng cửa sổ tài liệu.

---

## 📥 Hướng dẫn cài đặt nhanh (An toàn & Minh bạch 100%)

Bạn **không cần cài đặt Node.js**, không cần chạy dev server. Bạn chỉ cần tải dự án và chạy script cài đặt chính thức của Microsoft Office:

### Cách 1: Chạy file `install.ps1` bằng PowerShell (Khuyên dùng)
1. Tải toàn bộ mã nguồn dự án về máy (bấm nút xanh **Code** -> **Download ZIP** trên GitHub, rồi giải nén ra).
2. Nhấp chuột phải vào file **`install.ps1`** -> Chọn **Run with PowerShell** (Chạy bằng PowerShell).
3. Nếu hệ thống hỏi quyền Administrator (UAC), bấm **Yes**.
4. Script sẽ tự động thiết lập thư mục cài đặt `C:\AutoOfficeAddin` và đăng ký Trusted Catalog an toàn vào Microsoft Office.

*(Hoặc mở PowerShell với quyền Administrator tại thư mục và gõ: `.\install.ps1`)*

---

### Bước 2: Kích hoạt AutoOffice lần đầu trong Word
1. Mở **Microsoft Word** (mở một văn bản bất kỳ).
2. Trên thanh công cụ **Home**, bấm vào nút **Add-ins** -> Chọn **More Add-ins** *(hoặc vào thẻ **Insert -> Get Add-ins / My Add-ins**)*.
3. Trong hộp thoại **Office Add-ins**:
   - Chọn tab **SHARED FOLDER (Thư mục dùng chung)** ở phía trên.
   - Bấm chọn **AutoOffice** -> Nhấn **Add (Thêm)**.
4. Biểu tượng **AutoOffice** sẽ xuất hiện cố định ở góc bên phải của thẻ **Home**. Từ những lần mở Word tiếp theo, nút **AutoOffice** luôn nằm sẵn trên thanh công cụ!

---

### Bước 3: Nhập Khóa API Cá Nhân (Chỉ nhập 1 lần trên máy của bạn)
1. Bấm vào nút **AutoOffice** trên thanh **Home** để mở khung chat bên phải.
2. Ngay tại màn hình chào mừng (hoặc trong phần **⚙️ Cài đặt**), chọn nhà cung cấp AI:
   - **OpenAI** (`gpt-4o`, `gpt-5.4`, ...)
   - **Google Gemini** (`gemini-2.5-pro`, `gemini-2.5-flash`, ...)
   - **Anthropic Claude** (`claude-sonnet-4-6`, ...)
   - **DeepSeek** (`deepseek-chat`, ...)
   - Hoặc mô hình miễn phí chạy cục bộ qua **Ollama / LM Studio**.
3. Dán **API Key cá nhân** của bạn vào ô và bấm **Lưu Khóa API Cá Nhân**.

> 🔒 **Bảo mật & Riêng tư tuyệt đối**: Khóa API và lịch sử trò chuyện của bạn chỉ được lưu cục bộ trên máy tính cá nhân của bạn (`localStorage`), hoàn toàn không chia sẻ hay đồng bộ sang máy của người dùng khác.

---

## ✨ Điểm nổi bật của AutoOffice

1. **Bộ 10 nhóm tính năng thông minh cho Word (52 tác vụ chuyên sâu)**:
   - **Đọc nội dung & thống kê**: Đọc toàn bộ văn bản, nhận diện vùng đang bôi đen/bôi xanh, chữ in đậm/tô màu, bảng biểu, đếm số từ và dàn ý tiêu đề.
   - **Soạn thảo & sửa vùng bôi chọn**: Soạn văn bản hành chính đầy đủ Quốc hiệu Tiêu ngữ, sửa trực tiếp phần đang bôi đen/bôi đậm/tô màu, tìm & thay thế thông minh (không phân biệt hoa/thường, tự bắt qua nhiều dòng).
   - **Chuẩn hóa thể thức hành chính**: Tự động căn chuẩn văn bản hành chính Việt Nam (phông Times New Roman 13pt, giãn dòng 1.3, Quốc hiệu/Tiêu ngữ căn giữa, nội dung căn đều 2 bên).
   - **Bảng biểu & định dạng trang**: Tạo bảng biểu chuyên nghiệp, danh sách đánh số/gạch đầu dòng, căn lề trang A4 chuẩn Nghị định 30, chèn Header/Footer.
   - **Trường tự động & liên kết**: Chèn ngày tháng/số trang tự động, siêu liên kết (Hyperlink), ô điền biểu mẫu và hình ảnh minh họa.
   - **Nhận xét & Theo dõi sửa đổi**: Thêm/xóa/trả lời nhận xét bên lề trang, bật/tắt chế độ theo dõi sửa đổi (Track Changes).
   - **Mục lục tự động & Chú thích**: Chèn chú thích chân trang (Footnote/Endnote), đánh dấu trang và tạo bảng Mục lục tự động (TOC).
   - **Điền thông tin biểu mẫu**: Tự động điền hàng loạt chỗ trống "..." hoặc {{...}} trong mẫu đơn, chèn chữ chìm (Watermark).
   - **Thao tác nhanh**: Chèn tại vị trí con trỏ, cuộn tới đoạn văn, lưu tài liệu và phân tích mức độ dễ đọc.
   - **Giao diện thuần Việt**: Hoàn toàn ẩn mã lệnh kỹ thuật và JSON, chỉ hiển thị trạng thái tiếng Việt ngắn gọn và trả lời tự nhiên.
2. **Tự động chuẩn hóa Quốc hiệu & Tiêu ngữ**:
   - Khi yêu cầu tạo đơn từ, biên bản, báo cáo, bản kiểm điểm..., AutoOffice tự động đảm bảo có đầy đủ **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM** và **Độc lập - Tự do - Hạnh phúc** căn giữa chuẩn thể thức mà không làm rò rỉ gạch chân xuống các dòng dưới.
3. **Đa luồng độc lập cho từng cửa sổ Word**:
   - Mở nhiều file Word cùng lúc sẽ tạo các phiên trò chuyện độc lập (`📄 Tên_file.docx`), không bị xung đột dữ liệu giữa các cửa sổ và lưu trữ đầy đủ vào **Lịch sử**.

---

## 💻 Dành cho Lập trình viên (Tự chạy từ mã nguồn)

Nếu bạn muốn tự chạy máy chủ phát triển cục bộ:

```powershell
# 1. Cài đặt thư viện
npm install

# 2. Cài chứng chỉ HTTPS cục bộ (chọn Yes khi được hỏi)
npm run certs

# 3. Chạy server dev và tự động mở Word kèm Add-in
npm start

# 4. Hoặc chỉ bật server giao diện tại https://localhost:3721
npm run dev

# 5. Đóng gói bản dựng giao diện tĩnh (dist/)
npm run build
```
