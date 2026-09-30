# Demo giao diện Trợ lý dữ liệu quyết toán

Mở `windows.html` để xem **7 cửa sổ mẫu** và bấm thử các lựa chọn trong từng cửa sổ. Mở `index.html` để xem trang Thao tác cùng luồng đi vào các cửa sổ. Demo chạy tại máy, không cần cài thư viện, không đọc hoặc ghi dữ liệu của ứng dụng.

Trong PowerShell tại thư mục dự án:

```powershell
Start-Process .\demo-ui\windows.html
```

## Đường đi nên thử

1. Ở trang **Thao tác**, bước 1 vẫn mở Trợ lý ảo hoặc trợ lý bảng kê. Bước 2 mở **Xem file bóc tách**. Tại màn kiểm tra, lọc **Cần sửa**, mở khoản của hóa đơn `000849`, bổ sung ngày hóa đơn rồi lưu. Nút **Lưu và xác nhận** sẽ khả dụng.
2. Quay lại **Thao tác**. Ở bước 3, ba tác vụ **Hàng ngày → BK**, **Khoản chi → BK**, **BK → Thanh toán** vẫn độc lập. Nhấn **Nhập vào BK**, chọn tháng và phân tích để xem màn xử lý hai khoản chênh lệch. Không có file Excel nào được ghi.
3. Quay lại **Thao tác**. Ở bước 4, nhấn **Nhập PM quyết toán**, chọn sheet BK rồi chọn các số quyết toán. Demo không chạy RPA/PAD.
4. Các mục chính là **Thao tác**, **Lịch sử**, **Cài đặt**. Người dùng không thấy nhật ký kỹ thuật.

## Nguyên tắc thiết kế

- Tác vụ, trạng thái và bước tiếp theo hiển thị bằng ngôn ngữ nghiệp vụ.
- Giữ bốn nhóm thao tác và thứ tự của phần mềm hiện tại. Màn kiểm tra dùng danh sách bên trái và chi tiết bên phải; màn đối chiếu trình bày giá trị cũ và mới cạnh nhau.
- Hành động ghi hoặc xác nhận có bước xem trước. Trạng thái **đã khởi chạy** không được gọi là **đã nhập thành công**.
- Thông tin kỹ thuật nằm ở phần chi tiết hoặc thiết lập, vẫn có thể truy cập khi cần hỗ trợ.
- Bảng ưu tiên các cột dùng để ra quyết định. Thông tin nguồn bổ sung hiện khi chọn dòng.

## Đưa thiết kế vào ứng dụng Windows

Ứng dụng hiện tại dùng PySide6 và Qt Widgets. Có thể chuyển từng màn hình sang mẫu này bằng cách tạo các thành phần giao diện dùng chung (thẻ trạng thái, thanh tác vụ, panel chi tiết, hộp xác nhận) rồi gắn vào model và service hiện có. Mã nghiệp vụ, quy tắc kiểm tra, đối soát và ghi Excel cần được giữ nguyên khi thay giao diện.

HTML/CSS/JavaScript ở đây chỉ là nguyên mẫu để duyệt trải nghiệm. Nếu sau này cần hiệu ứng và chuyển cảnh nhiều hơn, Qt Quick/QML là một hướng cần đánh giá riêng vì sẽ tốn công chuyển các bảng và hộp thoại hiện có. Không cần thay framework để thực hiện các cải tiến trong bản demo.
