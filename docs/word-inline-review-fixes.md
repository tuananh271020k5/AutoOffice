# Accept / Reject trong Word và sửa lỗi vùng chọn, cỡ chữ

Luồng này thay cho việc mở một file khôi phục riêng.

## Duyệt ngay trên tài liệu

Khi bật **Duyệt chỉnh sửa Word bằng Accept / Reject**, các thao tác sửa nội dung,
định dạng và bảng trong thân tài liệu được hiển thị trực tiếp trong Word.
Khung chat chỉ hiện hai nút **Accept** và **Reject**, cùng phần chi tiết có thể mở rộng.

- **Accept** giữ thay đổi đang hiển thị và đọc lại để xác minh kết quả.
- **Reject** trả lại thân tài liệu trước thao tác ngay trong cùng file và chặn các
  thao tác ghi tiếp theo của lượt yêu cầu đó. Không tạo hoặc mở một file Word mới.
- Mỗi thao tác có quyết định riêng. Nếu AI tách một yêu cầu thành nhiều thao tác,
  ví dụ thay chữ rồi đổi font, có thể cần duyệt nhiều lần.
- Với header/footer, lưu file, nhận xét và các thao tác ngoài phạm vi thân tài liệu,
  vẫn hỏi duyệt trước khi chạy. Những thao tác này chưa có bản nháp tại chỗ để Reject.

Bản trước thao tác được giữ dưới dạng OOXML trong bộ nhớ. Trước khi Reject, AutoOffice
so sánh tài liệu hiện tại với trạng thái ngay sau thao tác. Nếu có thay đổi mới gây
xung đột, không ghi đè chúng và báo người dùng kiểm tra hoặc dùng Undo của Word.
Các bộ đếm, dấu hiệu kiểm tra chính tả và mã chỉnh sửa tự sinh của Word được loại
khỏi so sánh; chữ, định dạng, bảng, style và nội dung liên quan vẫn được so sánh.

Giới hạn: 8 MB XML cho bản trước thao tác; cần tắt Track Changes và xử lý các revision
đang có trước khi dùng chế độ này. Đây không phải khóa cộng tác: tránh sửa đồng thời
trong lúc AI đang thực thi. Bản nháp không lưu qua lần đóng/reload add-in. Nếu tắt
chế độ duyệt trong Cài đặt, các thay đổi chạy theo cấu hình tự duyệt cũ.

## Vùng chọn và chữ in đậm

Mục tiêu được cập nhật giữa các công cụ của cùng yêu cầu. Ví dụ, sau khi đổi `a`
thành `A`, bước đổi font kế tiếp nhắm vào `A`, không định dạng toàn tài liệu.
Khi toàn bộ yêu cầu kết thúc, con trỏ thu về cuối vùng chọn, cache và nhãn vùng chọn
được xóa. Các lượt đọc bất đồng bộ cũ không được ghi lại mục tiêu đã sử dụng.

Thay văn bản sẽ bỏ in đậm mặc định khi nội dung thực sự đổi, trừ khi yêu cầu nêu rõ
`bold: true`. Khi dùng đoạn in đậm làm dấu để tìm và thay thế, dấu đậm cũng được
tiêu thụ một lần. Lệnh chỉ yêu cầu in đậm vẫn giữ chữ in đậm theo yêu cầu.

## Cỡ chữ và xác minh

Các dạng `font_size`, `fontSize`, `size` được chuẩn hóa, kể cả khi nằm trong
`formatting`. Thao tác thay chữ áp dụng cả cỡ chữ được yêu cầu. Thay đổi định dạng
chỉ cập nhật các thuộc tính được chỉ định; không tự đổi font sang Times New Roman
hoặc cỡ 13 khi người dùng chỉ yêu cầu tăng cỡ chữ/in đậm.

Sau khi con trỏ chuyển đi, phép xác minh có thể tìm lại đúng văn bản mới để đọc font
của vùng đó, thay vì suy từ font của cả đoạn có nhiều định dạng. Nếu không xác định
được phạm vi duy nhất, kết quả vẫn được báo chưa xác minh, không khẳng định thành công.

## Kiểm thử local

```powershell
npm run test:word -- preview-test SESSION_ID reject
npm run test:word -- font-test SESSION_ID accept
npm run test:word -- font-test SESSION_ID format
npm run test:word -- font-test SESSION_ID format-reject
```

Các lệnh tạo dữ liệu trong tài liệu thử, gọi agent thật rồi chờ bấm Accept/Reject
trong Word. Báo cáo lưu trong `.autooffice-test/`. Bộ test tự động kiểm tra thêm
xung đột khi Reject, chuẩn hóa tham số, giữ font không được yêu cầu đổi, vùng chọn
một lần và xác minh vùng chữ sau khi con trỏ đã thu gọn.
