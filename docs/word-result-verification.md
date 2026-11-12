# Tính năng 1: đọc lại và xác minh kết quả trên Word

Các công cụ Word tích hợp sẵn nay quan sát tài liệu trước khi chỉnh sửa và đọc lại
bằng một `Word.run` mới sau khi thao tác thành công. Kết quả đối chiếu được đưa về
agent trong tool result để AI đánh giá tiếp yêu cầu, đồng thời lưu cùng thông báo
thao tác trong lịch sử chat. Không chỉ dựa vào thông báo `status: ok` của thao tác.

## Người dùng thấy gì?

- **Đang kiểm tra tài liệu…**: thao tác đã chạy, đang đọc lại kết quả.
- **Đã kiểm tra**: các điều kiện đo được của thao tác khớp với tham số yêu cầu.
- **Kết quả chưa khớp**: quan sát không khớp điều kiện đã kiểm tra.
- **Chưa xác minh**: không đọc được trước/sau, không xác định được vùng duy nhất,
  hoặc thao tác/thuộc tính chưa có bộ đối chiếu. Tài liệu có thể đã thay đổi.

Hai trạng thái cuối có cảnh báo kiểm tra tài liệu trước khi chạy lại. Chúng không
được hiển thị như một thao tác đã được xác nhận hoàn tất. Agent được hướng dẫn
không khẳng định thành công và không chạy lại thao tác chèn/xóa/thay thế một cách
mù quáng. Lớp kiểm tra không tự thực hiện lại hay tự hoàn tác thao tác.

## Các điều kiện hiện đối chiếu tự động

| Thao tác | Điều kiện đo được |
| --- | --- |
| Thay/xóa vùng chọn | Nội dung sau thao tác khớp phép thay thế trong toàn văn bản; yêu cầu vùng có nội dung duy nhất |
| Sửa đoạn văn theo chỉ số | Đúng nội dung đoạn được chỉ định; chữ của các đoạn khác giữ nguyên |
| Định dạng vùng chọn / danh sách chỉ số đoạn | Chữ không đổi; font, cỡ chữ, đậm, nghiêng, gạch chân, màu chữ và highlight được yêu cầu khớp với đọc lại |
| Tạo bảng với dữ liệu chỉ định | Có đúng một bảng mới, bảng cũ giữ nguyên, các hàng/cột và giá trị ô khớp |

Các thuộc tính không đo như bố cục trang, căn lề, style bảng hay thuộc tính định
dạng ngoài danh sách trên không được chứng nhận. Một thao tác có cả thuộc tính
đo được và thuộc tính chưa hỗ trợ sẽ trả **Chưa xác minh** khi phát hiện phạm vi
chưa được đo. Chứng nhận một phần nội dung không thay thế đánh giá toàn bộ yêu
cầu bằng ngôn ngữ tự nhiên của người dùng hoặc soát bố cục bằng mắt.

Các thao tác Word ghi khác vẫn được đọc lại nhưng trả **Chưa xác minh**, kèm bằng
chứng để AI xem xét. Các thao tác chỉ đọc không chạy thêm hai lần quan sát. Lưu
tài liệu không có API hỗ trợ nay báo lỗi thay vì trả `saved: true`.

## Giới hạn và chi phí

- Có thêm hai lần đọc Office.js cho mỗi thao tác ghi. Bằng chứng gửi lại model được
  giới hạn 12.000 ký tự; nếu bị cắt, agent phải đọc phần cụ thể khi cần.
- Các công cụ Word tích hợp được thực hiện tuần tự trong mỗi lượt agent để tránh
  kết quả của hai thao tác song song lẫn vào nhau. Đây không phải khóa tài liệu:
  người dùng, cộng tác viên hoặc MCP bên ngoài vẫn có thể chỉnh sửa đồng thời.
- Chưa áp dụng đối chiếu tự động cho công cụ MCP bên ngoài, Excel và PowerPoint.
- Chưa có nút hoàn tác hay hủy thực thi; đó là các tính năng tiếp theo.
- Nội dung snapshot được gửi đến model đang cấu hình giống ngữ cảnh tài liệu hiện
  có. Nó không chứa cấu hình/API key của người dùng.

## Kiểm tra

`npm test` kiểm tra cả đường chạy tích hợp: đọc trước/ghi/đọc sau, báo thành công sai,
lỗi đọc lại, thiếu baseline, thao tác bị từ chối, công cụ chỉ đọc, các thao tác đồng
thời và nhãn hiển thị cho người dùng.

Bộ `npm run test:word -- suite SESSION_ID` kiểm tra đầu-cuối trên Word thật qua cầu
nối local (xem [hướng dẫn](local-word-testing.md)). Báo cáo riêng cho từng lệnh có
`turn.toolResults[].result`: chuỗi JSON chứa `verification` và `readBack`.
