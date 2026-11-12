# Kiểm thử AutoOffice trên Word thật

Cầu nối này chỉ dành cho bản phát triển local. Nó gọi cùng `handleSend()` và
`runAgent()` như khung chat, sử dụng cấu hình model/API key hiện có trong add-in.
Không gửi API key cho người kiểm thử, không nhập key vào lệnh CLI.

## Khởi động

1. Dừng phiên sideload/dev server cũ bằng `npm run stop`. Nếu đã tự chạy Vite ở một
   terminal khác, dừng terminal đó để cổng 3721 trống.
2. Chạy `npm run sideload:test`. Lệnh này bật cầu nối rồi chạy luồng sideload hiện có.
3. Trong Word, mở khung AutoOffice và cấu hình provider/model/key nếu chưa có.
   Dùng tài liệu mẫu do sideload mở hoặc file tên `AutoOffice-Test*.docx`.
4. Giữ khung AutoOffice mở. Không đồng thời sửa tài liệu trong khi bộ thử đang chạy.
5. Chạy `npm run test:word -- sessions`, lấy đúng `sessionId` của cửa sổ thử.

```powershell
npm run test:word -- status SESSION_ID
npm run test:word -- snapshot SESSION_ID
npm run test:word -- run SESSION_ID "Tóm tắt tài liệu trong chat, không chỉnh sửa tài liệu."
npm run test:word -- suite SESSION_ID
```

`suite` **xóa nội dung tài liệu thử để tạo dữ liệu cho từng ca** và gọi AI thật,
có phát sinh token/chi phí theo nhà cung cấp. Bộ thử yêu cầu chế độ tự động duyệt
đang bật trong Cài đặt. Lệnh `run` riêng lẻ vẫn giữ nguyên chế độ duyệt của người
dùng; nếu cần duyệt, bấm duyệt trong Word. Cầu nối không tự đổi cài đặt này.

## Kết quả

- `.autooffice-test/report.json`: kết quả 4 ca: thay vùng chọn, tóm tắt không sửa,
  in đậm đúng đoạn và tạo bảng đúng ô.
- `.autooffice-test/COMMAND_ID.json`: kết quả từng lệnh, gồm snapshot trước/sau,
  lịch sử model, lỗi và kết quả công cụ ở lượt gửi.
- Có thể truy vấn lại lệnh bằng `npm run test:word -- result COMMAND_ID`.
- Đây là kiểm thử chức năng trên bộ tình huống cụ thể, không phải tỷ lệ chính xác
  tổng quát. Snapshot đo nội dung, định dạng đoạn và giá trị bảng; chưa kiểm tra
  bố cục bằng hình ảnh, mọi định dạng từng ký tự hay in/xuất PDF.

## Phạm vi và giới hạn

- Tắt mặc định; `npm run sideload` bình thường không bật cầu nối. Build production
  loại bỏ phần client. Không có API thực thi JavaScript tùy ý.
- Chỉ nhận kết nối loopback, Host localhost đúng cổng, Origin cùng server và token
  ngẫu nhiên mỗi lần chạy. Token nằm trong `.autooffice-test/connection.json`.
- Từng cửa sổ có hàng đợi riêng; lệnh bắt buộc chỉ định session. Client kiểm tra URL
  tài liệu thật trước mỗi lệnh và từ chối tài liệu không phải tài liệu thử.
- Cầu nối không xuất settings hay API key. Báo cáo có nội dung tài liệu và hội thoại;
  toàn bộ thư mục `.autooffice-test/` được gitignore.
- Không tự chạy lại thao tác khi mạng lỗi. Timeout chỉ dừng chờ, không hủy agent hoặc
  thao tác Word đang chạy. Kiểm tra trạng thái lệnh trước khi gửi tiếp; không reload
  add-in giữa lượt. Hàng đợi và kết quả trên server hết hạn sau 5 phút; file kết quả
  đã tải về vẫn còn. Khởi động lại server xóa hàng đợi đang lưu trong RAM.
- Khi xong chạy `npm run stop` rồi trở về `npm run sideload` để tắt chế độ thử.
