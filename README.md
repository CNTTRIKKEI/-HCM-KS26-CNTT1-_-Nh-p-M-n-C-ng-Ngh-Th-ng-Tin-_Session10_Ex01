# BỘ HỒ SƠ QUẢN LÝ DỰ ÁN VÍ ĐIỆN TỬ "PAYKX" (THÁNG 05/2024)

## PHẦN 1. PHÂN TÍCH CƠ SỞ DỮ LIỆU CỦA BẢNG GIAO DỊCH (EXCEL)

Trong mô hình dữ liệu chuẩn Database của bảng giao dịch:

- Record (Bản ghi / Hàng):

* Là một dòng dữ liệu duy nhất thể hiện thông tin chi tiết về một giao dịch phát sinh trên ứng dụng PayKX.

* Ví dụ: Bản ghi ở hàng 5 biểu thị toàn bộ thông tin đối soát của ngân hàng Vietcombank (STT: 1, Mã đối tác: VCB, Số giao dịch: 128,345, Doanh số: 25,364,580,000 VND,...).

- Field (Trường dữ liệu / Cột):

Là các thuộc tính mô tả đặc tính của giao dịch. Mỗi cột chứa một loại dữ liệu nhất định.

Các Field trong bảng dữ liệu bao gồm:

* `STT`: Số thứ tự bản ghi.

* `Ngân hàng đối tác`: Tên định danh ngân hàng liên kết.

* `Mã đối tác`: Mã khóa định danh ngắn (Key/Code).

* `Số giao dịch (thành công)`: Số lượng giao dịch hoàn tất.

* `Doanh số (VND)`: Tổng giá trị giao dịch.

* `Phí cố định (VND)`: Chi phí cố định thỏa thuận theo hợp đồng.

* `Tỷ lệ phí (%)`: Phần trăm phí biến đổi trên tổng doanh số.

* `phí biến đổi (VND)`: Chi phí phát sinh theo doanh số (=Doanh số+Tỷ lệ phí).

Tổng phí (VND): Tổng chi phí đối soát (=Phí cố định+phí biến đổi).

PHẦN 2. HƯỚNG DẪN TRIỂN KHAI BỘ 3 FILE (WORD, EXCEL, POWERPOINT)

1. FILE WORD: TÀI LIỆU MÔ TẢ TÍNH NĂNG HOÀN TIỀN (SRS)

A. Cấu trúc Heading Styles (Khung tài liệu 5 trang)

Để áp dụng Auto TOC (Mục lục tự động), cần gán chính xác các Heading Styles như sau:

Heading 1: Font Arial/Calibri, Size 16pt, Bold, Primary Color (Xanh lá/Xanh dương).

Heading 2: Font Arial/Calibri, Size 13pt, Bold, Dark Gray.

> 1. GIỚI THIỆU (Heading 1)
   1.1. Mục đích (Heading 2)
   1.2. Phạm vi (Heading 2)
> 2. TỔNG QUAN TÍNH NĂNG (Heading 1)
   2.1. Mô tả chung (Heading 2)
   2.2. Nguyên tắc hoạt động (Heading 2)
> 3. LUỒNG NGHIỆP VỤ (Heading 1)
   3.1. Điều kiện hoàn tiền (Heading 2)
   3.2. Quy trình hoàn tiền (Heading 2)
> 4. MÔ TẢ CHI TIẾT (Heading 1)
   4.1. Tạo yêu cầu hoàn tiền (Heading 2)
   4.2. Xử lý hoàn tiền (Heading 2)
   4.3. Thông báo kết quả (Heading 2)
> 5. QUY ĐỊNH & GIỚI HẠN (Heading 1)
   5.1. Giới hạn số tiền hoàn (Heading 2)
   5.2. Thời gian xử lý (Heading 2)
> 6. THÔNG BÁO & LỖI (Heading 1)
> 7. PHỤ LỤC (Heading 1)


B. Thao tác kỹ thuật trong Microsoft Word

Tạo Auto TOC: Vào tab References --> Chọn Table of Contents --> Chọn mẫu Automatic Table 1 hoặc Custom Table of Contents.

Phân trang đủ 5 trang: Sử dụng phím tắt Ctrl + Enter (Page Break) giữa các chương chính để đảm bảo tài liệu được phân bổ chuẩn mực trên 5 trang.

2. FILE EXCEL: BẢNG TỔNG HỢP PHÍ & ĐỐI SOÁT GIAO DỊCH

A. Cấu trúc Bảng dữ liệu chuẩn (Tuân thủ ràng buộc)

Tuyệt đối không Merge Cell trong vùng `$A\$4:\$H\$13$` (Bảng dữ liệu giao dịch).

Không để dòng trống giữa các bản ghi để đảm bảo tính toàn vẹn của Database.

B. Công thức chuẩn hóa và đối soát bằng VLOOKUP & TRIM

Giả định bảng tra cứu thông tin khách hàng nằm ở Sheet Dữ liệu giao dịch với phạm vi $A\$2:\$C\$100$ (Cột A: Mã giao dịch, Cột B: Mã người dùng, Cột C: Tên người dùng).

Làm sạch dữ liệu đầu vào bằng TRIM:
Để tránh lỗi do khoảng trắng thừa khi nhập liệu, công thức làm sạch mã giao dịch tại ô A2 là:

=TRIM(A2)


Công thức VLOOKUP kết hợp TRIM tra cứu Tên người dùng:

=VLOOKUP(TRIM(A2), 'Dữ liệu giao dịch'`!$A$2:$C$100`, 3, FALSE)


Giải thích tham số:

TRIM(A2): Giá trị tìm kiếm đã được loại bỏ khoảng trắng thừa ở đầu/cuối.

'Dữ liệu giao dịch'!$A$2:$C$100: Bảng chứa dữ liệu cần tra cứu (cố định bằng dấu $).

3: Cột chứa Tên người dùng (Cột C là cột thứ 3 trong vùng tra cứu).

FALSE: Tìm kiếm chính xác tuyệt đối.

3. FILE POWERPOINT: BÁO CÁO DOANH THU & LỖ LÃI (RULE 10/20/30)

A. Tiêu đề Slide (Assertion Title)

Sử dụng câu khẳng định truyền tải thông điệp chính thay vì tiêu đề chung chung:

"Doanh thu tăng 18,6% nhờ ưu đãi nạp điện thoại – Lợi nhuận sau thuế đạt 2,89 Tỷ VND"

B. Quy tắc thiết kế (Rule 10/20/30 & Ràng buộc)

Font size:

Tiêu đề slide (Assertion Title): 32pt - 36pt (Bold).

Chỉ số KPIs chính: 40pt - 48pt (Bold).

Nhãn biểu đồ / Nội dung giải thích: Tối thiểu 30pt (Theo yêu cầu đề bài).

Trực quan hóa dữ liệu (Thay thế bảng số liệu thô):

Biểu đồ Cột (Column Chart): Thể hiện sự tăng trưởng "Doanh thu thuần" qua 5 tháng (T01: 7,28B $\rightarrow$ T05: 12,45B).

Biểu đồ Tròn/Đồng xu (Donut Chart): Thể hiện "Cơ cấu chi phí" Tháng 05/2024 (Phí đối tác & ngân hàng: 46%, Chi phí vận hành: 28%, Chi phí tiếp thị: 16%, Khác: 10%).

PHẦN 3. KIỂM THỬ VÀ XỬ LÝ LỖI (QUY TRÌNH 4 ĐIỂM KHI VLOOKUP BỊ #N/A)

Khi hàm =VLOOKUP(TRIM(A2), Target_Range, Col_Index, FALSE) trả về kết quả #N/A, thực hiện quy trình kiểm tra 4 điểm theo các bước sau:

[Điểm 1: Khóa tìm kiếm] ──> [Điểm 2: Định dạng dữ liệu] ──> [Điểm 3: Vùng tra cứu] ──> [Điểm 4: Sự tồn tại dữ liệu]


Bước 1: Kiểm tra khoảng trắng và ký tự ẩn ở ô dữ liệu nguồn (Lookup Value)

Nguyên nhân: Dù đã dùng TRIM, các ký tự khoảng trắng không ngắt dòng (Non-breaking space, mã ASCII 160 - thường do copy từ web/system) vẫn không thể bị loại bỏ.

Cách khắc phục:

Sử dụng hàm CLEAN kết hợp hoặc dùng tính năng Find & Replace (Ctrl + H) để xóa ký tự khoảng trắng lạ.

Công thức nâng cao: =VLOOKUP(TRIM(SUBSTITUTE(A2, CHAR(160), "")), ...)

Bước 2: Kiểm tra sự đồng nhất về Kiểu dữ liệu (Data Type Mismatch)

Nguyên nhân: Mã giao dịch ở ô tìm kiếm đang ở dạng Text (Chuỗi), nhưng mã trong Bảng tra cứu lại ở dạng Number (Số), hoặc ngược lại.

Cách khắc phục:

Nếu chuyển từ Text sang Number: =VLOOKUP(VALUE(TRIM(A2)), ...)

Nếu chuyển từ Number sang Text: =VLOOKUP(TEXT(A2, "0"), ...)

Bước 3: Kiểm tra Vùng tra cứu (Table Array) & Dấu cố định địa chỉ ($)

Nguyên nhân:

Cột chứa khóa tìm kiếm không nằm ở cột đầu tiên (Cột thứ nhất) của vùng tra cứu Table_Array.

Quên cố định địa chỉ vùng tra cứu (dùng F4 để có dấu $), dẫn đến khi kéo công thức xuống các hàng dưới, vùng tra cứu bị lệch (Data drift).

Cách khắc phục:

Đảm bảo cột khóa nằm ở vị trí đầu tiên trong vùng được chọn.

Cố định vùng tra cứu thành dạng $A\$2:\$C\$100$.

Bước 4: Kiểm tra sự tồn tại thực sự của dữ liệu trong Bảng nguồn

Nguyên nhân: Mã giao dịch thực sự chưa tồn tại trong cơ sở dữ liệu đối soát của ngân hàng đối tác.

Cách khắc phục:

Dùng tính năng Ctrl + F tìm kiếm thủ công trong bảng gốc để xác nhận sự tồn tại.

Sử dụng hàm IFERROR để quản trị trải nghiệm hiển thị khi dữ liệu thực sự chưa cập nhật:

=IFERROR(VLOOKUP(TRIM(A2), 'Dữ liệu giao dịch'!$A$2:$C$100, 3, FALSE), "Chưa đối soát")
