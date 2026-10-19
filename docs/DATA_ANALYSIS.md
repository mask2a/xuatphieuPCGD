# Phân tích dữ liệu nguồn

## Phạm vi và cách đọc

Đã đọc workbook bằng Excel COM (Excel 16.0) ở chế độ chỉ đọc. Hai tệp yêu cầu
hiện nằm ở thư mục gốc workspace, không nằm trong `templates/`:

| Vai trò | Tệp | Sheet chính |
|---|---|---|
| Dữ liệu nguồn | `dulieumau.xls` | `MauNhapLieu` |
| Mẫu phiếu | `maubaocao.xls` | `Phiếu điều tra 1 mặt` |

`dulieumau.xls` còn có các sheet `DuLieu` và `HuongDan`; sheet `Sheet1` rỗng.
Sheet `MauNhapLieu` có UsedRange `A1:AY1059`; các cột có tiêu đề và dữ liệu là
`A:AX` (50 cột). Dữ liệu nguồn bắt đầu tại hàng 5. Hàng 1 là nhóm tiêu đề,
hàng 2 là tiêu đề cột, hàng 3 bổ sung một số tiêu đề con, hàng 4 đánh số cột.

## Kết quả đếm trên tệp mẫu

Chỉ tính các hàng từ 5 trở đi có ít nhất một giá trị trong `B:AY`:

| Chỉ số | Kết quả |
|---|---:|
| Bản ghi thành viên | 427 |
| Mã số phiếu `N` không rỗng, duy nhất | 120 |
| Bản ghi không có tên (cả `B` và `C` trống) | 0 |
| Bản ghi thiếu ngày / tháng / năm sinh | 0 / 0 / 1 |
| Bản ghi thiếu mã số phiếu | 0 |
| Mã phiếu có nhiều thành viên | 102 |
| Mã phiếu có nhiều tên chủ hộ khác nhau | 3 |
| Mã phiếu có nhiều địa chỉ khác nhau | 0 |
| Mã phiếu có nhiều điện thoại không rỗng khác nhau | 28 |
| Mã phiếu có nhiều giá trị diện cư trú khác nhau | 5 |

Phân bố số thành viên theo 120 mã phiếu: 1 người: 18 hộ; 2: 15; 3: 16;
4: 36; 5: 25; 6: 9; 8: 1. Mẫu này không có hộ 7 người hoặc trên 8 người.
Đây là thống kê của riêng tệp mẫu, không phải kết quả của ứng dụng.

Giá trị diện cư trú quan sát được ở cột `O`: `Thường trú`, `Tạm trú`,
`thường trú`, `Thường Trú`. Cột `P` (`Tình trạng cư trú`) có giá trị
`Lưu Trú`; không được đồng nhất cột `P` với cột `O`.

Có 28 mã phiếu có nhiều số điện thoại khi xét tất cả thành viên. Theo quy tắc
đã chốt, đây **không** tự động là mâu thuẫn: chỉ dùng số `AW` của bản ghi có
`AU` xác định duy nhất là chủ hộ. Không lấy số của thành viên khác để thay
thế. Nếu chủ hộ không có số, để trống; nếu không xác định duy nhất chủ hộ
hoặc có nhiều bản ghi chủ hộ với số khác nhau, để trống và ghi cảnh báo.
Ba mã phiếu có biến thể tên chủ hộ trong cột `K:L` cần được cảnh báo. Năm mã
phiếu có nhiều giá trị diện cư trú cần được cảnh báo. Không sửa dữ liệu gốc
khi chuẩn hóa hay xuất.

## Danh mục cột trên `MauNhapLieu`

Cột là ký hiệu Excel; tên được đọc từ hàng 2, hoặc hàng 3 khi hàng 2 là tiêu
đề nhóm. `—` nghĩa là không có trường riêng trên phiếu đầu ra hoặc chưa đủ cơ
sở để xác định ô đích; xem bảng ánh xạ để biết lý do.

| Cột | Tiêu đề nguồn | Đích / ghi chú |
|---|---|---|
| A | TT | Số thứ tự trên mẫu là số vị trí thành viên, không phải số thứ tự toàn tệp; không chép đè lên số vị trí |
| B | Họ đệm | Kết hợp với `C` vào ô họ tên thành viên |
| C | Tên | Kết hợp với `B` vào ô họ tên thành viên |
| D | Ngày | Thành phần ngày sinh; định dạng nguồn `00` |
| E | Tháng | Thành phần tháng sinh; định dạng nguồn `00` |
| F | Năm sinh | Thành phần năm sinh |
| G | Nữ | `X` nếu nữ; trống không có nghĩa là được phép suy ra giới tính khác |
| H | Dân tộc | Ô dân tộc trên phiếu |
| I | Tôn giáo | Ô tôn giáo; rỗng trong cả 427 bản ghi mẫu |
| J | Diện ưu tiên | Không có ô tương ứng rõ ràng |
| K | Chủ hộ | Phần họ đệm tên chủ hộ; ghép với `L` |
| L | Tên | Phần tên chủ hộ |
| M | Địa chỉ: Số nhà, tên đường, tổ (nếu có) | Ô địa chỉ/Xóm; mẫu đầu ra không có trường địa chỉ đầy đủ riêng |
| N | Số phiếu | Mã nhóm hộ; dữ liệu dạng chuỗi, cần giữ nguyên |
| O | Diện cư trú | Ô diện cư trú; khác cột `P` |
| P | Tình trạng cư trú | Cảnh báo/đối chiếu với cột AF trên phiếu |
| Q | Khối học | Thông tin học hiện tại; phiếu có các dòng năm học cố định |
| R | Lớp học | Thông tin lớp hiện tại; nguồn không có năm học tương ứng |
| S | Quận/Huyện đang học | Chỉ có 1 giá trị không rỗng trong tệp mẫu |
| T | Tên trường đang học | Thông tin trường; phiếu chỉ có một ô trường-huyện mỗi dòng |
| U | Mã Trường | Không có ô mã trường riêng trên phiếu |
| V | Bậc tốt nghiệp | Ứng viên ở "Cấp học" trong khối tốt nghiệp; cần quy tắc chọn dòng |
| W | Bổ túc | Ứng viên cột "Bổ túc"; cần quy tắc chọn dòng |
| X | Năm tốt nghiệp | Ứng viên cột "Năm học"; cần quy tắc chọn dòng |
| Y | Bậc TN Nghề | Ứng viên cột "Bậc" TN nghề; cần quy tắc chọn dòng |
| Z | Số bằng tốt nghiệp | Không có ô số bằng riêng trên phiếu |
| AA | Năm TN nghề | Ứng viên cột "Năm học" TN nghề; cần quy tắc chọn dòng |
| AB | Học xong | Ứng viên cột "Học xong" |
| AC | Năm | Ứng viên năm học hoàn thành |
| AD | Bỏ học | Ứng viên cột "Bỏ học" |
| AE | Năm | Ứng viên năm bỏ học |
| AF | Đang học lớp | Ứng viên cột "Đang học lớp"; cần xác nhận ý nghĩa với `Q/R` |
| AG | Hoàn thành lớp | Ứng viên cột "Hoàn thành lớp" |
| AH | Tái mù chữ mức | Ô mức tái mù chữ |
| AI | Khuyết tật vận động | Các loại khuyết tật nguồn cùng dùng một cột "Loại" trên phiếu |
| AJ | Khuyết tật nghe, nói | Như trên |
| AK | Khuyết tật nhìn | Như trên |
| AL | Khuyết tật thần kinh, tâm thần | Như trên |
| AM | Khuyết tật trí tuệ | Như trên |
| AN | Khuyết tật học tập | Như trên |
| AO | Tự kỷ | Như trên |
| AP | Khuyết tật khác | Như trên |
| AQ | Có chứng nhận khuyết tật | Cột "Có CNKT" |
| AR | Khả năng học tập | Cột "KNHT" |
| AS | Hoàn cảnh đặc biệt | Không có trường tương ứng rõ ràng |
| AT | Chi tiết hoàn cảnh đặt biệt | Không có trường tương ứng rõ ràng |
| AU | Quan hệ với chủ hộ | Ô quan hệ của thành viên |
| AV | Họ tên Cha hoặc Mẹ | Ô cha/mẹ/người đỡ đầu |
| AW | Điện thoại | Chỉ dùng số của thành viên được xác định duy nhất là chủ hộ qua `AU`; không lấy số thành viên khác thay thế |
| AX | Ghi chú | Ô ghi chú của thành viên |

`D` và `E` là giá trị số có định dạng hiển thị `00`; ví dụ giá trị lưu `1`
hiển thị `01`. Khi đọc, không được làm mất định dạng ngày/tháng. `N` và điện
thoại trong mẫu là chuỗi; không ép sang số. Ô trống phải được giữ trống, không
tự điền ngày 01, tôn giáo, giới tính hay giá trị khác.

## Hạn chế và điểm cần xác nhận

1. Cột `N` có thể lặp lại giữa các tệp. Chế độ thư mục phải phát hiện trùng mã
   giữa tệp và không gộp chỉ dựa trên `N`. Ngay trong tệp mẫu này, ba mã có tên
   chủ hộ khác nhau; cần cảnh báo thay vì tự chọn một giá trị.
2. Đã chốt quy tắc điện thoại: chỉ lấy `AW` của thành viên duy nhất có quan
   hệ chủ hộ ở `AU`; nếu không xác định duy nhất hoặc các bản ghi chủ hộ có
   số khác nhau thì để trống và cảnh báo. Các số điện thoại khác của thành
   viên không tạo cảnh báo xung đột.
3. Diện cư trú là thông tin cấp hộ; các mã có nhiều giá trị ở cột `O` cần
   được cảnh báo, không tự chọn giá trị.
4. Tệp nguồn có một bản ghi thiếu năm sinh. Giữ trống và báo cảnh báo.
5. Tệp mẫu nằm ở thư mục gốc, không có thư mục `templates/`. Workspace còn có
   tệp `12..10 - PHỔ CẬP TDP 18 (1).xls`; tệp này không được yêu cầu trong
   phạm vi phân tích giai đoạn 1 và chưa được dùng để tính thống kê.
