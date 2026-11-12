# Tính năng 2: xem trước và mở bản khôi phục

> Luồng mở bản khôi phục riêng trong tài liệu này đã được thay bằng
> [Accept / Reject ngay trong Word](word-inline-review-fixes.md). Phần dưới lưu lại
> thiết kế cũ để tham khảo lịch sử, không mô tả giao diện hiện tại.

Trong Cài đặt có **Xem trước và giữ bản khôi phục (Word)**, bật mặc định kể cả
với cấu hình cũ. Khi bật, chỉnh sửa bằng công cụ Word tích hợp luôn chờ người dùng
duyệt, dù **Tự động thực thi chỉnh sửa** đang bật. Công cụ chỉ đọc không hỏi duyệt.

## Xem trước

Trước thao tác, AutoOffice đọc Word để trình bày phạm vi, nội dung hiện tại,
nội dung dự kiến và các thuộc tính định dạng phổ biến. Dữ liệu bảng có bản xem
trước các ô. Người dùng chọn **Áp dụng thay đổi** hoặc **Giữ nguyên tài liệu**.

Đây là bản mô tả thao tác dự kiến, chưa phải bản render Word sau chỉnh sửa.
Định dạng tự động, một số thao tác nâng cao và một số thuộc tính không có bản
xem trước chính xác. Nội dung vượt 16.000 ký tự và bảng vượt 20 hàng hoặc 10 cột
được rút gọn kèm thông báo. Với thao tác chưa hỗ trợ xem trước chính xác, UI báo
rõ giới hạn và phạm vi tài liệu.

Sau khi duyệt, chữ, thuộc tính font đọc được, dữ liệu bảng và vùng chọn được đọc
lại. Nếu các dữ liệu đang quan sát khác trạng thái đã duyệt,
thao tác bị từ chối và AI phải đọc lại để đề xuất bản mới. Kiểm tra này cũng chạy
sau quá trình tạo bản khôi phục. Nó không khóa được các chỉnh sửa đồng thời của
người dùng/cộng tác viên trong khoảng thời gian giữa hai lần gọi Word.

Vùng chọn gốc được giữ bằng tracked range trong thời gian duyệt. Khi việc bấm vào
khung duyệt làm Word thu gọn vùng chọn, phạm vi gốc được chọn lại trước khi kiểm
tra và chỉnh sửa. Cách này giữ đúng vị trí kể cả văn bản trùng nhau; tham chiếu được
giải phóng khi thao tác kết thúc hoặc bị từ chối.

Nếu người dùng từ chối, các thao tác ghi Word tích hợp tiếp theo trong cùng lượt
agent bị chặn, kể cả AI thử đổi sang công cụ Word khác.

## Khôi phục

Trước thao tác ghi đầu tiên đã duyệt trong mỗi yêu cầu, AutoOffice dùng
`Office.context.document.getFileAsync(Office.FileType.Compressed)` để giữ một bản
DOCX đầy đủ. Những thao tác tiếp theo của cùng yêu cầu dùng chung bản trước lượt
đó. Nếu không tạo được bản sao đầy đủ, thao tác được bảo vệ không chạy.

Sau đó khung chat có nút **Mở bản khôi phục**. Nút này dùng
`Word.Application.createDocument(base64).open()` để **mở thành tài liệu riêng**.
Tài liệu đang chỉnh sửa được giữ lại, gồm cả những thay đổi người dùng vừa làm.
Người dùng so sánh và Save As bản muốn giữ. Đây không phải nút Undo trong tài liệu
gốc, và không làm thay đổi file gốc trên ổ đĩa.

- Chỉ giữ bản sao gần nhất trong bộ nhớ của cửa sổ add-in. Yêu cầu chỉnh sửa mới
  sẽ thay bản sao cũ sau khi tạo bản mới thành công.
- Bản sao không lưu vào lịch sử chat, localStorage, máy chủ kiểm thử hoặc gửi cho
  model. Đóng/reload khung AutoOffice làm mất bản sao; hãy lưu bản khôi phục nếu cần.
- Giới hạn 20 MB cho tài liệu nén. Word phải hỗ trợ các API đọc file và mở tài liệu
  từ Base64; quyền truy cập file/chính sách tổ chức cũng có thể hạn chế chúng.
- Có thể tắt chế độ này trong Cài đặt để dùng luồng cũ. Khi tắt, không tạo bản khôi
  phục mới; cài đặt tự duyệt trước đó lại quyết định có hỏi duyệt thao tác ghi không.
- Phạm vi: công cụ Word tích hợp. MCP bên ngoài, Excel, PowerPoint chưa đi qua luồng
  xem trước/bản khôi phục này.

## Kiểm thử

`npm test` kiểm tra preview, phạm vi vùng chọn, mặc định định dạng, từ chối duyệt,
vùng chọn/tài liệu thay đổi sau duyệt, lỗi backup, đọc file nhiều slice, timeout,
đóng handle, mở bản riêng, nhãn/nút UI và chuyển đổi cấu hình cũ.

Với cầu nối local, `run` đi qua đúng luồng duyệt và tạo bản sao. Hãy duyệt/từ chối
trong Word. Bộ `suite` tự động yêu cầu tắt chế độ này một cách chủ động vì không
tự bấm duyệt thay người dùng. Không tắt bảo vệ chỉ để tránh chờ mà chưa xem nội dung
tài liệu thử.

`npm run test:word -- preview-test SESSION_ID reject` và `... accept` tạo dữ liệu
thử, gửi yêu cầu thật rồi chờ quyết định qua các nút trong Word. Kết quả được lưu
ở `.autooffice-test/preview-reject.json` và `preview-accept.json`. Sau ca chấp nhận,
dùng nút **Mở bản khôi phục** để kiểm tra bản trước chỉnh sửa.
