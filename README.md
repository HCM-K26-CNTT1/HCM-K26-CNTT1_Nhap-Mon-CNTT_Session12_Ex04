# BÀI 4: CHUẨN HÓA QUY TRÌNH QUẢN LÝ BỆNH ÁN ĐIỆN TỬ

## 1. Giới thiệu dự án

**Tên hệ thống:** Clinic-Cloud

**Bối cảnh:** Phòng khám Startup cần hiện đại hóa việc lưu trữ và quản lý hồ sơ bệnh nhân bằng hệ thống điện tử.

**Mục tiêu:**
- Xây dựng tài liệu SRS cho module Bệnh án điện tử.
- Quản lý 1.000 hồ sơ bệnh nhân bằng Excel theo Database Mindset.
- Trực quan hóa số liệu theo nhóm bệnh bằng biểu đồ Donut Chart.
- Đảm bảo an toàn, bảo mật và tính toàn vẹn của dữ liệu bệnh nhân.

## 2. Cấu trúc bộ hồ sơ

```text
Clinic-Cloud/
├── README.md
├── SRS.docx
├── DanhSachBenhNhan.xlsx
└── BaoCaoDichBenh.pptx
```

## 3. Phần 1. Phân tích

### 3.1. Yêu cầu chức năng (Functional Requirements)

Yêu cầu chức năng mô tả những công việc hệ thống cần thực hiện để hỗ trợ quy trình khám bệnh.

- Đăng ký bệnh nhân và tạo mã bệnh nhân duy nhất.
- Tra cứu hồ sơ theo mã bệnh nhân hoặc họ tên.
- Cho phép bác sĩ cập nhật thông tin khám bệnh, chẩn đoán và hướng điều trị theo quyền được cấp.
- Quản lý tài khoản và phân quyền người dùng.

### 3.2. Yêu cầu phi chức năng (Non-functional Requirements)

Yêu cầu phi chức năng mô tả các tiêu chuẩn về chất lượng và cách hệ thống phải vận hành.

- **Bảo mật:** Yêu cầu đăng nhập và phân quyền truy cập hồ sơ.
- **Toàn vẹn dữ liệu:** Hạn chế mã bệnh nhân trùng lặp và kiểm tra dữ liệu trước khi lưu.
- **Sao lưu:** Có phương án sao lưu và khôi phục khi xảy ra sự cố.
- **Hiệu năng:** Hỗ trợ quản lý ít nhất 1.000 hồ sơ và tra cứu nhanh trong điều kiện vận hành xác định.
- **Ghi nhận hoạt động:** Lưu lại các thao tác quan trọng để phục vụ kiểm tra khi cần.

**Kết luận:** Yêu cầu chức năng xác định hệ thống làm được những gì; yêu cầu phi chức năng xác định hệ thống phải đáp ứng các tiêu chuẩn như bảo mật, hiệu năng và độ tin cậy như thế nào.

## 4. Phần 2. Triển khai hệ thống Clinic-Cloud tích hợp

### 4.1. File Word – SRS Bệnh án điện tử

Tài liệu SRS gồm các nội dung:

- Giới thiệu và phạm vi hệ thống.
- Đối tượng sử dụng: bác sĩ, nhân viên tiếp nhận và quản trị viên.
- Yêu cầu chức năng quản lý hồ sơ và khám bệnh.
- Yêu cầu phi chức năng về bảo mật, sao lưu, hiệu năng và toàn vẹn dữ liệu.
- Quy tắc quản lý dữ liệu và phương án kiểm thử.

Sử dụng Heading 1 và Heading 2 để phân cấp tiêu đề, sau đó tạo mục lục tự động bằng chức năng Table of Contents trong Word.

### 4.2. File Excel – Quản lý 1.000 hồ sơ bệnh nhân

Bảng dữ liệu gồm các trường: Mã bệnh nhân, Họ và tên, Ngày sinh, Giới tính, Số điện thoại, Ngày khám và Nhóm bệnh.

Áp dụng Database Mindset:

- Mỗi hàng tương ứng với một bệnh nhân.
- Mỗi cột chỉ chứa một loại thông tin.
- Mã bệnh nhân phải duy nhất.
- Không gộp ô hoặc chèn dòng trống hoàn toàn giữa các bản ghi.
- Ô chưa có thông tin được để trống, không tự điền dữ liệu giả.
- Sử dụng Table và AutoFilter để quản lý, tìm kiếm và lọc hồ sơ.

**Xử lý bẫy dữ liệu:** Chọn toàn bộ vùng dữ liệu, nhấn `Ctrl + T`, tích chọn `My table has headers` rồi nhấn OK. Cách này giúp AutoFilter áp dụng cho toàn bộ bảng, kể cả khi một số ô thông tin bị bỏ trống. Không cần xóa các ô trống hợp lệ trong hồ sơ.

Tạo thêm Sheet thống kê số bệnh nhân theo nhóm bệnh để cung cấp số liệu cho biểu đồ PowerPoint.

### 4.3. File PowerPoint – Báo cáo số liệu dịch bệnh

Bài trình bày gồm 3 slide:

- **Slide 1 – Tổng quan:** Giới thiệu hệ thống Clinic-Cloud và mục tiêu quản lý hồ sơ.
- **Slide 2 – Thống kê nhóm bệnh:** Sử dụng Donut Chart để thể hiện tỷ lệ bệnh nhân theo từng nhóm bệnh.
- **Slide 3 – Giải pháp:** Trình bày các biện pháp bảo mật, sao lưu và chuẩn hóa dữ liệu.

Sử dụng thiết kế tối giản, nền sáng, màu sắc dễ phân biệt và cỡ chữ dễ đọc. Tiêu đề slide nên nêu kết luận chính thay vì chỉ ghi tên chủ đề.

*Lưu ý: Số liệu trên biểu đồ phải lấy từ bảng thống kê Excel hoặc được ghi rõ là số liệu giả định minh họa.*

## 5. Phần 3. Kiểm thử và xử lý lỗi

### 5.1. Lỗi gán nhầm Heading Styles

**Tình huống:** Một đoạn văn nội dung bị gán nhầm thành Heading 1, khiến đoạn văn xuất hiện trong mục lục tự động.

**Cách khắc phục:**

1. Bôi đen đoạn văn bị định dạng sai.
2. Nhấn `Ctrl + Shift + N` để chuyển đoạn văn về kiểu Normal.
3. Kiểm tra lại định dạng và cấu trúc tiêu đề.
4. Nếu mục lục chưa thay đổi, nhấp phải vào mục lục, chọn Update Field và cập nhật toàn bộ mục lục.

### 5.2. Kiểm thử dữ liệu Excel

- Kiểm tra mã bệnh nhân có bị trùng không.
- Lọc theo nhóm bệnh và kiểm tra kết quả.
- Kiểm tra các hồ sơ có ô thông tin trống.
- Đảm bảo AutoFilter hoạt động trên toàn bộ bảng dữ liệu.
- Không xóa hồ sơ chỉ vì một số trường thông tin chưa được cung cấp.

### 5.3. Kiểm thử bảo mật

- Người chưa đăng nhập không được truy cập hồ sơ.
- Nhân viên chỉ được xem hoặc sửa dữ liệu theo quyền hạn.
- Kiểm tra khả năng sao lưu và khôi phục dữ liệu.
- Kiểm tra việc ghi nhận các thao tác quan trọng trên hệ thống.

## 6. Kết luận

Hệ thống Clinic-Cloud kết hợp tài liệu SRS, bảng tính Excel và báo cáo PowerPoint để chuẩn hóa quy trình quản lý bệnh án điện tử. Việc phân biệt yêu cầu chức năng với phi chức năng giúp xác định rõ nhiệm vụ và tiêu chuẩn chất lượng của hệ thống. Excel Table giúp quản lý dữ liệu có ô trống mà vẫn duy trì khả năng lọc, trong khi Heading Styles và quy trình kiểm thử hỗ trợ tài liệu nhất quán, dễ sử dụng và dễ bảo trì.
