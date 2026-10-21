# Xác minh nguồn hiện tại — 2026-10-21

> Cập nhật mẫu mới nhất: [kiểm tra cấu trúc thực tế](TEMPLATE_STRUCTURE_AND_SAMPLES_REVIEW.md).
> Theo xác nhận người dùng, mẫu không còn phải khớp hash snapshot cũ; hash chỉ ghi log, cấu trúc sai mới chặn xuất.

## Chính sách hiện hành sau xác nhận của người dùng

Người dùng xác nhận dữ liệu `dulieumau.xls` nguyên vẹn và yêu cầu không
kiểm tra lại nội dung. Nguồn không còn phải khớp SHA-256, mtime, kích thước
hay đường dẫn trong snapshot cũ. Snapshot vẫn được giữ nguyên để tham khảo.
Test `test_source_has_required_structure_and_can_be_grouped` dùng reader
thật kiểm tra sheet, tiêu đề bắt buộc B/C/N và dữ liệu đọc được; kiểm tra
nhóm hộ giữ đầy đủ các bản ghi, không mất hoặc nhân đôi khi nhóm.
Reader hiện có tiếp tục từ chối workbook không có tiêu đề bắt buộc hợp lệ.

Template tiếp tục kiểm tra nghiêm ngặt hash/metadata qua
`test_template_matches_approved_snapshot_strictly`; các test COM cấu trúc
không thay đổi. Không bổ sung chức năng xuất hay thay đổi mapping.

Sau thay đổi, chạy `.\venv\Scripts\python.exe -m unittest discover -v`
trong phiên Windows với Tkinter/Excel COM: **38 PASS, 0 FAIL, 0 SKIP**
(285,242 giây). Không sửa hai workbook gốc hoặc snapshot; không xuất
hàng loạt và không commit.

## Kết quả xác minh trước xác nhận (lịch sử)

Trạng thái tại thời điểm xác minh: **CHỜ NGƯỜI DÙNG CHẤP THUẬN HASH NGUỒN MỚI**.
Kết quả chi tiết và toàn bộ tên/vị trí 50 cột A:AX tại hàng 2–3 nằm trong
`CURRENT_SOURCE_VERIFICATION.json`. Không có dữ liệu định danh thành viên
trong báo cáo này.

## Nguồn

- Excel COM mở `dulieumau.xls` với `ReadOnly=True`, `UpdateLinks=0`,
  tắt macro và sự kiện, đóng bằng `SaveChanges=False`.
- Sheet/vùng: `MauNhapLieu!A1:AY1059`, `Sheet1!A1` (rỗng),
  `DuLieu!A1:Z333`, `HuongDan!A1:D44`.
- COM và reader hiện tại cùng đọc được **427 nhân khẩu, 120 hộ**.
  Reader không báo lỗi. Dữ liệu bắt đầu hàng 5; tên cột lấy tại hàng 2–3.
- Nhóm hộ giữ 18 cảnh báo: 10 không xác định được chủ hộ duy nhất,
  5 xung đột diện cư trú, 3 xung đột tên chủ hộ.
- Các cột phục vụ giá trị xuất: B:C (họ tên), D:F (ngày sinh), H:I
  (dân tộc/tôn giáo), K:L (chủ hộ), N (mã hộ), R/T/V/X
  (lớp/trường/bậc và năm tốt nghiệp), AU:AW (quan hệ, cha/mẹ, điện thoại).
  Q dùng đối chiếu lớp, không ghi. Xã và năm học lấy từ cấu hình.
  Đối chiếu giá trị COM với reader tại 18 cột này trên 427 bản ghi:
  **0 khác biệt**. Đây là đối chiếu hai cách đọc phiên bản hiện tại,
  không phải đối chiếu với dữ liệu phiên bản cũ.
- Các trường chưa duyệt vẫn không được ghi; mapping không thay đổi.
- Kích thước 623104 bytes, bằng snapshot. Mtime UTC hiện tại
  `2026-10-20T16:58:54+00:00`, snapshot `2026-10-19T14:25:01+00:00`.
- SHA-256 hiện tại:
  `a9cdc8672b65f57ce00a1a2b36940f406fdd45f8c70574d3706692df3d629b58`.
- SHA-256 đã duyệt:
  `e7ce0b359bc516d54463a29ab88142ff355b735bda6408f77566dc5e0d3b262a`.

## Đối chiếu bản cũ

Không tìm thấy workbook có SHA-256 nguồn đã duyệt trong các tệp `.xls`
hiện có trong dự án. Tệp `12..10 - PHỔ CẬP TDP 18 (1).xls` có hash khác,
không được coi là bản snapshot cũ. Git không có lịch sử `dulieumau.xls`.
Snapshot JSON chỉ chứa metadata/hash, không chứa giá trị ô.
Vì vậy **không thể khẳng định dữ liệu không thay đổi hoặc liệt kê các ô
thay đổi so với phiên bản đã duyệt**. Số lượng bản ghi/hộ trùng chưa chứng
minh nội dung từng ô trùng.

## Template

- `maubaocao.xls`: 54272 bytes, mtime UTC `2026-10-18T12:52:27+00:00`.
- SHA-256 `963a75a4ae9258ff40f45006f97a9750c31023e0cee7bb4c115d938c52a9434e`
  khớp snapshot.
- Excel COM chỉ đọc xác minh sheet `Phiếu điều tra 1 mặt`, vùng A1:AG94,
  14 vị trí thành viên, 442 vùng gộp. Các kiểm tra cấu trúc của exporter
  `_verify_template_sheet` và `_template_signature` hoàn tất: nhãn cố định,
  thứ tự thành viên, ô gộp, ô đích giáo dục đã duyệt, định dạng,
  kích thước hàng/cột và thiết lập trang đọc được.
- Hash và metadata cả hai workbook trước/sau kiểm tra không đổi.

## Tách kiểm tra và quy trình chấp thuận

- `test_template_matches_approved_snapshot_strictly`: kiểm tra nghiêm ngặt
  path, size, mtime và SHA-256 template. Các test COM cấu trúc hiện có giữ nguyên.
- `test_source_requires_content_verification_and_user_approval_when_hash_changes`:
  SHA-256 là điều kiện bắt buộc, mtime chỉ để theo dõi trong báo cáo.
  Hash khác gây FAIL với thông báo phải xác minh nội dung và lấy chấp thuận
  rõ ràng của người dùng trước khi dùng phiên bản mới. Không tự ghi snapshot.
- Xác minh kỹ thuật không phải chấp thuận. Chưa sử dụng nguồn mới để xuất
  phiếu; cần người dùng chấp thuận chính xác hash mới trước khi cập nhật
  snapshot qua một thay đổi riêng được cho phép.
- Thay đổi hiện tại tách các test bảo vệ; không bổ sung cổng hash trong
  GUI/exporter. Không được hiểu rằng GUI đã có cơ chế chấp thuận phiên bản.

## Kiểm chứng

Chạy các test analysis/report, grouper, reader và hai test snapshot mới:
**15 PASS, 1 FAIL, 0 SKIP** (16 test). FAIL là nguồn chưa được chấp thuận,
đúng cơ chế bảo vệ; template PASS. Không chạy lại toàn bộ suite trong lượt này.
Kiểm tra COM chỉ đọc hoàn tất riêng, không xuất workbook hoặc PDF.
Không sửa workbook, mapping, cấu hình thật hay snapshot; không commit.
