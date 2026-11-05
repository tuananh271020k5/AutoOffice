# 🚀 AutoOffice — Trợ lý AI thông minh cho Microsoft Word, Excel & PowerPoint

**AutoOffice** là tiện ích (Add-in) tích hợp trực tiếp vào thanh công cụ của **Microsoft Word, Excel và PowerPoint**, trang bị **Bộ 10 Công cụ MCP chuyên sâu (52 tác vụ)** giúp soạn thảo văn bản hành chính Việt Nam chuẩn Nghị định 30/2020/NĐ-CP, nhận diện trực tiếp vùng văn bản đang bôi chọn/in đậm/tô màu, và chạy đa luồng độc lập trên từng cửa sổ tài liệu.

---

## 📥 Tải về bộ cài đặt nhanh (Dành cho người dùng Word)

Bạn **không cần cài đặt Node.js, không cần tải mã nguồn và không cần gõ lệnh**. Chỉ cần tải file cài đặt `.exe` (20 KB) dưới đây:

👉 **[⬇️ TẢI NGAY: AutoOffice-Setup.exe (GitHub Release v1.0.0)](https://github.com/anhkkk5/AutoOffice/releases/latest/download/AutoOffice-Setup.exe)**  
*(Link dự phòng trực tiếp: [Tải AutoOffice-Setup.exe](https://github.com/anhkkk5/AutoOffice/raw/master/releases/AutoOffice-Setup.exe))*

---

## 🛠️ Hướng dẫn cài đặt vào Microsoft Word (Chỉ mất 30 giây)

### Bước 1: Chạy bộ cài `AutoOffice-Setup.exe`
1. Nhấp đúp vào file **`AutoOffice-Setup.exe`** vừa tải về.
2. Nếu Windows hỏi quyền quản trị (UAC), bấm **Yes** (hoặc *More info -> Run anyway* nếu SmartScreen hỏi).
3. Trong cửa sổ hiện ra, bấm nút màu xanh **"Cài đặt AutoOffice"** -> Đợi thông báo **Cài đặt hoàn tất!**.

*(Bộ cài sẽ tự động đóng các cửa sổ Word đang mở, tạo thư mục chia sẻ cục bộ `C:\AutoOfficeAddin` và đăng ký Trusted Catalog vào Microsoft Office).*

---

### Bước 2: Kích hoạt AutoOffice lần đầu trong Word
1. Mở **Microsoft Word** (mở một tài liệu bất kỳ).
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

1. **Bộ 10 Công cụ MCP chuyên sâu cho Word (52 Actions/Ops)**:
   - `word_read`: Đọc cấu trúc tài liệu, đoạn văn `[P0]..[Pn]`, nhận diện vùng đang bôi chọn/in đậm/tô màu, đếm từ và dàn ý.
   - `word_edit`: Soạn văn bản, sửa trực tiếp phần đang bôi chọn (`edit_selection`), sửa theo định dạng in đậm/tô màu (`edit_by_format`), tìm kiếm & thay thế thông minh không phân biệt hoa/thường và khoảng cách (`search_and_replace`), hỗ trợ cả văn bản trong khung **Text Box** trang bìa.
   - `word_format`: Chuẩn hóa văn bản hành chính Việt Nam (`format_administrative_vn` — Times New Roman 13pt, giãn dòng 1.3, căn đều 2 bên), định dạng font, đoạn văn, kiểu chữ và tô sáng.
   - `word_structure`: Tạo và định dạng bảng biểu (`insert_table`), ngắt trang (`insert_break`), thiết lập khổ giấy A4 & lề chuẩn (`set_page_setup`), Header/Footer.
   - `word_media`: Chèn hình ảnh Base64, liên kết (Hyperlink) và đánh dấu (Bookmark).
   - `word_review`: Thêm/xóa nhận xét (Comments) và bật/tắt theo dõi sửa đổi (Track Changes).
   - `word_references`: Tạo mục lục tự động (TOC), chú thích chân trang (Footnote/Endnote) và trường dữ liệu (Fields).
   - `word_document`: Quản lý Content Controls, thuộc tính tài liệu và xuất/nhập OOXML.
   - `word_app`: Kiểm tra môi trường Office.js.
   - `word_help`: Tra cứu danh mục công cụ và hướng dẫn nhanh.
2. **Tự động chuẩn hóa Quốc hiệu & Tiêu ngữ**:
   - Khi yêu cầu tạo đơn từ, biên bản, báo cáo, bản kiểm điểm..., AutoOffice tự động đảm bảo có đầy đủ **CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM** và **Độc lập - Tự do - Hạnh phúc** căn giữa chuẩn thể thức mà không làm rò rỉ gạch chân xuống các dòng dưới.
3. **Đa luồng độc lập cho từng cửa sổ Word**:
   - Mở nhiều file Word cùng lúc sẽ tạo các phiên trò chuyện độc lập (`📄 Tên_file.docx`), không bị xung đột dữ liệu giữa các cửa sổ và lưu trữ đầy đủ vào **Lịch sử**.

---

## 💻 Dành cho Lập trình viên (Tự chạy từ mã nguồn)

Nếu bạn muốn tự chạy máy chủ phát triển cục bộ hoặc tự đóng gói lại bộ cài `.exe`:

### Yêu cầu:
- **Node.js >= 22** (Khuyên dùng bản LTS tại [nodejs.org](https://nodejs.org/))
- **Microsoft Office** (Word / Excel / PowerPoint trên Windows)

### Các lệnh chính:
```powershell
# 1. Cài đặt thư viện
npm install

# 2. Cài chứng chỉ HTTPS cục bộ (chỉ cần chạy 1 lần đầu, chọn Yes khi được hỏi)
npm run certs

# 3. Chạy server dev và tự động mở Word kèm Add-in
npm start

# 4. Hoặc chỉ bật server giao diện tại https://localhost:3721
npm run dev

# 5. Đóng gói bản dựng giao diện tĩnh (dist/)
npm run build

# 6. Đóng gói lại file cài đặt installer/output/AutoOffice-Setup.exe
npm run build:exe
```

---

## 🗑️ Gỡ cài đặt
Nếu muốn gỡ bỏ Add-in khỏi Word, bạn có thể mở lại file `AutoOffice-Setup.exe` và bấm nút **"Gỡ cài đặt"**, hoặc chạy lệnh:
```powershell
AutoOffice-Setup.exe /UNINSTALL /S
```
