# Mapping được duyệt — 2026-10-21

> Cập nhật tiếp theo: [kiểm tra cấu trúc và đối chiếu 6 mẫu](TEMPLATE_STRUCTURE_AND_SAMPLES_REVIEW.md).
> Người dùng đã xác nhận mẫu được lưu lại; hash chỉ theo dõi, cấu trúc mới là điều kiện chặn xuất.

Quy tắc mới thay thế các phần “chưa duyệt” tương ứng trong tài liệu cũ.
Không thay đổi workbook gốc, snapshot mẫu hoặc mở khóa Xuất tất cả/PDF.

## Vị trí và hành vi

- **Xóm (2)** thực tế ở A3:G3, không phải hàng 2. Ghi M + xã/phường từ
  GUI, không lặp địa danh hoặc dấu phẩy. Nếu địa chỉ các thành viên mâu
  thuẫn: để trống, cảnh báo, không tự chọn.
- **Diện cư trú:** Z2:AG2 có hai checkbox rich text. Chỉ một lựa chọn
  được thay bằng X; lựa chọn kia để trống. Trống/khác/mâu thuẫn: không
  đánh dấu, có cảnh báo. Chỉ font dấu X đổi sang Arial để hiển thị đúng.
- **Nữ:** C tại dòng thứ tư (C12 cho người đầu); G có X hoặc Nữ thì ghi X,
  Nam/trống để trống. Font ô X đổi từ Wingdings sang Arial có chủ đích.
  Dân tộc D12 và tôn giáo E12 giữ logic cũ.
- **Năm:** mọi khối dùng 2025-2026 đến 2029-2030, không đổi năm vùng ký.
  GUI khởi tạo 2025-2026; dù cấu hình lưu năm khác, exporter luôn ghi năm
  cố định. Interface cập nhật nhãn linh hoạt độc lập vẫn được giữ nguyên.
- **Lớp/trường/cấp học/năm tốt nghiệp:** giữ H/K/P/R hiện có; chỉ dòng
  năm thứ nhất. Không suy ra cấp học. Bổ túc chưa có logic ghi trong mã cũ;
  giữ nguyên theo mục 6 của yêu cầu, không tự bổ sung W → Q.

Các vị trí sau được xác minh qua Excel COM chỉ đọc; reader lưu cả tiêu đề
cha và con sau khi giải gộp tiêu đề. Resolver tìm đúng tên, không đoán cột.
Tiêu đề thiếu hoặc trùng: cảnh báo SOURCE_HEADER_UNRESOLVED và không ghi.

| Nguồn đã xác minh | Tiêu đề thực | Đích trên dòng năm thứ nhất |
|---|---|---|
| Y | Bậc TN Nghề | S |
| AA | Năm TN nghề | T:U |
| AB | Học xong / Lớp | V |
| AC | Học xong / Năm | W |
| AD | Bỏ học / Lớp | X |
| AE | Bỏ học / Năm | Y |
| AF | Đang học lớp | Z |
| AG | Hoàn thành lớp | AA |
| AH | Tái mù chữ mức | AB |
| AI:AP | Nhóm tiêu đề khuyết tật | AC |
| AQ | Có chứng nhận khuyết tật | AD |
| AR | Khả năng học tập | AE |
| AT | Chi tiết hoàn cảnh đặt biệt | AF |
| AX | Ghi chú | AG |

“đặt biệt” là cách viết thực trong tiêu đề AT; không sửa nguồn. Loại
khuyết tật chỉ chọn các ô X, ghép tên loại bằng dấu phẩy + khoảng trắng,
bỏ tiền tố “Khuyết tật”. Không có dấu X: để trống. Nếu thiếu/trùng một
tiêu đề loại, không ghi AC và cảnh báo. Mọi giá trị khác giữ nguyên nguồn.

Trước khi xuất kiểm tra tất cả 14 bộ ô đích mới: đúng merge, không có công
thức hoặc dữ liệu sẵn. Sau lưu mở lại kiểm tra giá trị mới của từng người.
Các trường ngoài yêu cầu vẫn để trạng thái chưa ánh xạ.

## Bố cục

- Hộ 1–7: hiện cả 7 khối đầu, chỉ ẩn từ khối 8 đến trước phần ký.
- Hộ 8–14: hiện đủ 14 khối. Không xóa hàng hay khối trống.
- Mỗi sheet in A1:AG87 và FitToPages 1×1.
- Trên 14: giữ cơ chế **sheet bổ sung** và đánh số liên tục của mã cũ;
  mã hiện có không chèn khối vào cùng sheet.
- Giữ 442 vùng gộp, viền, kích thước hàng/cột, khổ/lề trang và vùng ký.

## Kiểm chứng và tệp

Mã: `reader.py`, `gui.py`, `src/excel_exporter.py`,
`src/approved_mapping.py`, `config/template_mapping.json`.
Test: `tests/test_approved_mapping.py`, `tests/test_excel_exporter.py`.
Xuất sáu hộ tổng hợp qua `scripts/export_mapping_samples.py`; script không
đọc thành viên thật, không xuất 120 hộ.

Kết quả xuất thử tại `output/testing/approved_mapping_20261021_010939/`:

| Workbook | Thành viên | Số trang thực tế Excel GET.DOCUMENT(50) |
|---|---:|---|
| ho_SYNTHETIC-01.xlsx | 1 | 1 |
| ho_SYNTHETIC-05.xlsx | 5 | 1 |
| ho_SYNTHETIC-07.xlsx | 7 | 1 |
| ho_SYNTHETIC-08.xlsx | 8 | 1 |
| ho_SYNTHETIC-14.xlsx | 14 | 1 |
| ho_SYNTHETIC-15.xlsx | 15 | 1 + 1 trên sheet bổ sung |

`VERIFICATION.json` ghi kết quả và xác nhận hash nguồn/mẫu/snapshot trước
và sau xuất thử không đổi. Giá trị các trường mới của từng người được
exporter mở lại để so sánh trước khi tính thành công. Đây là thử nghiệm
kỹ thuật với dữ liệu tổng hợp, không phải nghiệm thu nghiệp vụ.

Lưu ý mẫu: hash quan sát được trong lượt này là
`fd4211e34c95d4de267267860e20279e50ea8a84f06f4682018894115efec4d1`,
khác snapshot `963a75a4ae9258ff40f45006f97a9750c31023e0cee7bb4c115d938c52a9434e`.
Kích thước và mtime vẫn khớp. Không tự sửa snapshot hoặc bỏ test bảo vệ;
không kết luận nguyên nhân thay đổi chỉ từ hash.

## Kết quả cuối

Lệnh `.\venv\Scripts\python.exe -m unittest discover -v`:
**42 PASS, 1 FAIL, 0 SKIP, 0 ERROR** (43 test, 314,435 giây).
FAIL duy nhất là `test_template_matches_approved_snapshot_strictly` do
hash mẫu nêu trên. Các test mapping mới, cấu trúc COM, GUI, reader,
grouper và bảo vệ công thức đều PASS. Không thay đổi snapshot để làm test PASS.
`py_compile` và `git diff --check` trên các tệp thay đổi đều PASS.

Các tiêu đề mới trên 427 bản ghi nguồn hiện tại không có trường hợp
thiếu/trùng; không có trường mới nào phải bỏ ghi do chưa xác định được
tiêu đề. Những tệp khác trong workspace đã có thay đổi trước lượt này
được giữ nguyên, không hoàn tác. Không commit, không xuất hàng loạt,
không tuyên bố nghiệm thu. Cần người dùng kiểm tra GUI và xác nhận phiên
bản mẫu hiện tại trước khi thay đổi snapshot trong một yêu cầu riêng.
