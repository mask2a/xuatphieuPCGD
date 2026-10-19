# Báo cáo audit cuối Giai đoạn 3 — Xuất Excel

Ngày kiểm tra: 2026-10-19

## Kết luận

**Chưa nghiệm thu chính thức Giai đoạn 3.** Bộ xuất Excel hoạt động trên
mẫu thật và dữ liệu tổng hợp, nhưng cấu hình địa bàn/năm học chưa được cung
cấp hợp lệ và nhiều trường nghiệp vụ còn ánh xạ chưa được xác minh. Không
xuất phiếu chính thức từ 120 hộ nguồn; không triển khai PDF hoặc GUI.

## Độ phủ ánh xạ

| Chỉ số | Kết quả |
|---|---:|
| Quy tắc đích đã xác minh trong mapping JSON | 10 |
| Cột nguồn được bao phủ bởi các quy tắc trên | 13 |
| Trường nghiệp vụ khai báo chưa ánh xạ | 36 |
| Trường mới được phép bổ sung vào vùng ghi | 0 |

Mười quy tắc đã có gồm 4 trường cố định đầu phiếu và 6 trường trong khối
thành viên. Các cột nguồn B, C, D, E, F, H, I, K, L, N, AU, AV, AW được
bao phủ. A là số thứ tự dòng nguồn, không được chép sang số thứ tự thành
viên trên phiếu.

Không có ứng viên mới nào đủ cả tọa độ và quy tắc nghiệp vụ để được ghi:
giới tính G và diện cư trú O chưa chốt ký hiệu checkbox; địa chỉ M không
đồng nhất với nhãn “Xóm”; các nhóm trường học/tốt nghiệp cần quy tắc năm;
các trường khuyết tật cần cách ghép giá trị; J, U, Z, AS, AT không có ô
đích đã xác minh. Cột J đã được đưa vào danh sách cảnh báo chưa ánh xạ.
Chi tiết từng trường/15 lần cảnh báo ở [MAPPING_AUDIT.md](./MAPPING_AUDIT.md).

## Phân tích nguồn thật

Đọc chỉ-đọc `dulieumau.xls`: **1 tệp đọc, 0 lỗi, 427 bản ghi, 427 thành
viên, 120 hộ**. Phân bố hộ có 1/2/3/4/5/6/8 thành viên lần lượt là
18/15/16/36/25/9/1. Không có hộ 7 hoặc trên 8 người trong nguồn.

Các mâu thuẫn được giữ nguyên, không tự chọn:

- 3 hộ có tên chủ hộ không thống nhất.
- 5 hộ có diện cư trú không thống nhất.
- 10 hộ không xác định được duy nhất chủ hộ; điện thoại hộ để trống theo
  quy tắc hiện hành. Điện thoại của người khác không được dùng thay.

Đã xác định được đại diện theo kích thước 1, 6 và 8 thành viên cùng các hộ
có mâu thuẫn để dùng cho lần kiểm tra thật tiếp theo. **Không xuất các hộ
này:** `config/settings.json` vẫn có `commune_name` rỗng và `school_year`
null. Người dùng trả lời “đồng ý” ở các trường giá trị không cung cấp tên xã
hay năm học; hai câu đó không được ghi vào cấu hình và không được xem là dữ
liệu hợp lệ.

## Kiểm thử Excel

Lệnh đã chạy:

```text
.\venv\Scripts\python.exe -m unittest discover -v
```

Kết quả: **21 PASS, 0 FAIL**. Mười chín kiểm thử trước vẫn PASS; thêm hai
kiểm thử preflight cấu hình:

1. Thiếu xã/năm học thì xuất bị từ chối trước khi tạo thư mục đầu ra.
2. Năm học sai định dạng hoặc địa bàn hộ khác cấu hình thì xuất bị từ chối.

Kiểm thử Excel COM trên Windows vẫn mở lại workbook xuất, không báo Repair,
và xác minh cấu trúc/ô đã ánh xạ. Năm workbook tổng hợp 1, 6, 7, 14, 15
người trước đó đạt **5 thành công, 0 thất bại**, có 15 cảnh báo về các
trường chưa ánh xạ. Đây không phải là kết quả xuất từ hộ thật.

Test kiểm tra cảnh báo cho diện ưu tiên J khi có dữ liệu; dữ liệu mẫu hiện
không có giá trị J khác rỗng. Quy tắc điện thoại chủ hộ và các mâu thuẫn
tên/diện cư trú tiếp tục được kiểm thử, không sửa dữ liệu nguồn.

## Cấu hình và an toàn khi xuất

Exporter đọc `config/settings.json` mặc định hoặc đường dẫn cấu hình truyền
rõ. Trước khi tạo thư mục đầu ra/khởi chạy Excel COM, yêu cầu:

- `commune_name` không rỗng;
- `school_year` dạng `YYYY-YYYY`, hai năm liên tiếp;
- `commune_name` của mọi hộ khớp cấu hình.

Giá trị năm học được đưa vào nội dung cảnh báo nhưng chưa ghi lên năm in sẵn
trong mẫu: chưa có quy tắc xác định các cột/hàng nguồn phải ứng với từng
năm. Vì vậy cấu hình hợp lệ vẫn chưa đủ để nghiệm thu nghiệp vụ.

## Số hộ xuất thử

| Nhóm | Thành công | Thất bại | Kết luận |
|---|---:|---:|---|
| Tổng hợp trên mẫu thật, dữ liệu kiểm thử tổng hợp (1/6/7/14/15 người) | 5 | 0 | PASS kỹ thuật; không phải dữ liệu hộ thật |
| Hộ từ `dulieumau.xls` | 0 | 0 | Không chạy do cấu hình bắt buộc chưa hợp lệ; không tính là PASS hay FAIL |

Hash nguồn/mẫu được giữ nguyên trong kiểm thử trước; suite hiện tại cũng chỉ
đọc dữ liệu nguồn và mở mẫu chỉ đọc. Không phát hành workbook hộ thật.

## Điều kiện còn thiếu để nghiệm thu

1. Cung cấp **tên xã thật** cho `commune_name` và **năm học thật** theo định
   dạng `YYYY-YYYY`; không dùng chuỗi “đồng ý” hoặc giá trị mẫu.
2. Xác nhận ký hiệu đánh dấu nữ và diện cư trú trên mẫu (G, O).
3. Xác nhận cách hiển thị địa chỉ M tại trường “Xóm”.
4. Xác nhận quan hệ từng cột học tập/tốt nghiệp với các hàng năm trên phiếu.
5. Xác nhận cách biểu diễn các loại khuyết tật, trường hợp nhiều loại, và
   hàng đặt ghi chú AX.
6. Xác nhận requiredness nghiệp vụ cho các trường mà nguồn có thể để trống.
7. Sau khi chốt, bổ sung mapping/test cho từng quy tắc, xuất hộ thật đại
   diện gồm hộ mâu thuẫn để xác minh không tự chọn, rồi đối chiếu đủ 120 hộ
   trước khi đánh dấu PASS chính thức.

Không có thay đổi dữ liệu nguồn hoặc mẫu. Chặng PDF và GUI vẫn chưa bắt đầu.
