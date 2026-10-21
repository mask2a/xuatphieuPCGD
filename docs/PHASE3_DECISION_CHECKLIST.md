# Checklist quyết định trước nghiệm thu Giai đoạn 3

Ngày lập: 2026-10-20

## 1. Những vấn đề đã xác minh

- Cấu hình nằm trong `config/settings.json`; `settings.py` có
  `load_settings(path)` để nạp cấu hình và
  `validate_export_settings(settings)` để từ chối thiếu xã/năm học hoặc năm
  học không theo `YYYY-YYYY` với hai năm liên tiếp.
- Cấu hình hiện tại là `commune_name=null`,
  `school_year="2026-2027"`. Loader không lấy địa danh in sẵn trên mẫu làm
  mặc định; `null` được nạp thành chuỗi rỗng và sẽ chặn xuất chính thức.
- `settings.py` có `save_settings(...)` để cập nhật hai khóa này bằng cách
  ghi tệp tạm cùng thư mục rồi thay thế nguyên tử; các khóa cấu hình khác
  được giữ nguyên. GUI chưa được triển khai.
- Trong cấu trúc hiện tại không có `src/settings.py`. Module cấu hình ở
  thư mục gốc là `settings.py`.
- Mẫu `maubaocao.xls` đọc chỉ-đọc bằng Excel COM:
  - `B12` là nhãn “Nữ:”; `C12` có `q`, font `Wingdings`.
  - `Z2:AG2` có nội dung “Diện cư trú (5): q Thường trú q Tạm trú”.
  - `F9:F13` có các năm in sẵn 2025–2026 đến 2029–2030.
  - Các ô nhập quan sát tại `H9:J9`, `K9:O9`, `P9`, `AC9:AG9` đang trống.
  - Không có trường/nhãn diện ưu tiên riêng quan sát được trên phiếu.
- Hash mẫu trước/sau kiểm tra hiện tại không đổi:
  `89354830a91694e6a9251ccd0b02214c19625c211da683e1a086a74411694aad`.
- 36 trường chưa ánh xạ được phân loại **A=0, B=31, C=5**. Không có
  mapping mới được thêm.
- Quy tắc chủ hộ tiếp tục giữ nguyên: chỉ dùng điện thoại của chủ hộ duy
  nhất; không lấy số thành viên khác thay thế; mâu thuẫn tên/diện cư trú
  không được tự lựa chọn.

## 2. Những vấn đề chưa xác minh

- Chưa có bản sao/bản kiểm soát cũ của `dulieumau.xls` để đối chiếu digest
  `65b9...1f10` với digest hiện tại `5382...81c0`.
- Báo cáo JSON cũ ghi tên `dulieumau.xls` và hai giá trị SHA-256, nhưng
  không ghi đường dẫn tuyệt đối, kích thước, mtime hay quy trình hash byte
  cụ thể. Không thể khẳng định chỉ từ tên file rằng hai digest đại diện
  cùng một tệp tại cùng một đường dẫn.
- Chưa xác định nguyên nhân hoặc thời điểm hash khác nhau; không suy luận
  rằng file đã bị sửa bởi người hay quy trình cụ thể.
- Ký tự `q` ở `C12` và trong `Z2` chưa được kiểm chứng bằng tài liệu quy
  ước hoặc phiếu điền chuẩn để biết trạng thái chọn/không chọn được render
  ra sao. Giá trị COM chỉ phản ánh nội dung/format lưu trong workbook.
- Chưa xác định quy tắc nhập giới tính G, diện cư trú O, địa chỉ M vào
  nhãn “Xóm”, tình trạng cư trú P, hoặc ánh xạ học tập/tốt nghiệp theo năm.
- Chưa xác định requiredness pháp lý/nghiệp vụ cho 31 trường B. Các mức ưu
  tiên đề xuất trong [MAPPING_AUDIT.md](./MAPPING_AUDIT.md) không thay thế
  phê duyệt nghiệp vụ.
- Giao diện nhập/lưu cấu hình chưa tồn tại. Hàm lưu mới cung cấp seam để
  GUI tương lai cập nhật xã/năm học; giao diện vẫn phải nhập/validate hai
  giá trị mà không hard-code.

## 3. Các quyết định bắt buộc trước khi xuất thật

- [ ] Người dùng xem xét sự khác nhau giữa hash nguồn cũ/mới và xác nhận
  phiên bản `dulieumau.xls` nào được dùng làm đầu vào có thẩm quyền. Nếu
  cần đối chiếu, cung cấp bản nguồn cũ hoặc manifest đáng tin cậy; không
  khôi phục/thay thế tự động.
- [ ] Cung cấp tên xã thật cho `commune_name` và xác nhận
  `school_year="2026-2027"` là đúng; lưu/xác nhận trong
  `config/settings.json`. Không xuất chính thức khi xã còn `null`.
- [ ] Trả lời nhóm câu hỏi B1–B15 trong
  [USER_DECISIONS.md](./USER_DECISIONS.md) (đối chiếu bảng bằng chứng trong
  [MAPPING_AUDIT.md](./MAPPING_AUDIT.md)), hoặc phê duyệt rõ trường nào được
  giữ trống/cảnh báo trong phạm vi xuất được chấp thuận. Không xóa cảnh báo
  khi chưa có quy tắc.
- [ ] Quyết định cho 5 trường nhóm C (J, U, Z, AS, AT): xác nhận chúng nằm
  ngoài mẫu/được phép bỏ khỏi phiếu này, hay cung cấp biểu mẫu/đặc tả khác.
  Không nhét các dữ liệu này vào ô khác.
- [ ] Xác nhận requiredness: trường nào phải có dữ liệu, trường nào được
  trống khi nguồn thiếu; không suy diễn từ thực tế cột đang trống.
- [ ] Duyệt quy trình xử lý hộ mâu thuẫn, bảo đảm hộ không bị tự chọn tên
  chủ hộ/diện cư trú và điện thoại chỉ lấy từ chủ hộ duy nhất.
- [ ] Sau các phê duyệt trên, cho phép tạo một lượt xuất thử hộ thật đại
  diện; trước khi chạy, ghi lại path/size/mtime/SHA-256 của cả nguồn và
  mẫu, sau khi chạy xác nhận hash không đổi.

## 4. Những quyết định có thể hoãn

- GUI, PDF, gộp PDF, đóng gói EXE và thao tác nhập/lưu cấu hình qua GUI có
  thể hoãn; chúng không thuộc bước xác minh hiện tại.
- Các ánh xạ chỉ áp dụng cho hồ sơ không có dữ liệu trong bộ nguồn được
  duyệt có thể hoãn nếu người dùng chấp thuận rõ việc tiếp tục để trống và
  giữ cảnh báo. Không được dùng sự chấp thuận đó để gọi là ánh xạ đầy đủ.
- Các trường nhóm C chỉ có thể hoãn nếu người dùng xác nhận chấp nhận phạm
  vi mẫu hiện tại không biểu diễn chúng; nếu chúng bắt buộc cho phiếu thì
  phải đổi mẫu/đặc tả trước nghiệm thu.

## 5. Kết quả đối chiếu hash

| Thuộc tính | Nguồn hiện tại | Báo cáo đối chiếu cũ |
|---|---|---|
| Tên/định danh | `dulieumau.xls` | Ghi `dulieumau.xls` trong mô tả dữ liệu tổng hợp |
| Đường dẫn tuyệt đối | `C:\Users\Admin\Desktop\VSCODE PROJECT\PHOCAPGD\dulieumau.xls` | Không có |
| Kích thước | 623,104 bytes | Không có |
| Mtime | 2026-10-19 14:25:01.044 UTC (2026-10-19 21:25:01.044 theo giờ máy) | Không có |
| SHA-256 | `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0` | `65b9aa56e10bfaaa13e0d19d10eebcc09622484c126132677c46711121961f10` trước/sau |
| Thuật toán được thể hiện | SHA-256 trên bytes của file hiện tại | Các khóa JSON ghi `source_sha256_*`; không ghi đường dẫn hay phương pháp thực thi |

Hai digest khác nhau. Cùng basename và tên thuật toán không đủ để chứng
minh cùng một đối tượng tệp; báo cáo cũ thiếu metadata nhận dạng. Không có
bản nguồn cũ được kiểm soát để so sánh và không thể xác định nguyên nhân
chênh lệch. Trong lần đọc hiện tại, SHA-256 trước/sau giữ nguyên; không
chỉnh sửa nguồn hoặc mẫu. Vì người dùng yêu cầu xem xét chênh lệch trước
khi xuất dữ liệu thật, **xuất hộ thật vẫn bị dừng**.

## 6. Điều kiện đề xuất nghiệm thu Giai đoạn 3

1. Xác định và phê duyệt phiên bản nguồn; lưu manifest kiểm thử gồm đường
   dẫn tuyệt đối, kích thước, mtime UTC và SHA-256 trước/sau.
2. Có giá trị cấu hình xã/năm học hợp lệ; kiểm thử từ chối thiếu/sai cấu
   hình trước tạo output/khởi chạy COM.
3. Có quyết định bằng văn bản cho B1–B15 và nhóm C; mapping JSON chỉ thay
   đổi cho các quy tắc đã có bằng chứng, kèm test đối chiếu độc lập nguồn →
   ô đích.
4. Chạy Excel COM trên hộ thật đại diện (hộ nhỏ/lớn trong nguồn, thiếu dữ
   liệu và từng loại mâu thuẫn) sau khi người dùng cho phép; xác minh số
   người, thứ tự, trường, số điện thoại chủ hộ, merge, vùng in, phần xác
   nhận, định dạng và mở lại workbook.
5. Xác nhận mọi cảnh báo được giải thích; không mất dữ liệu/không tự chọn
   giá trị mâu thuẫn; hash nguồn/mẫu không đổi.
6. Mới sau đó đối chiếu kết quả của toàn bộ 120 hộ và quyết định PASS/FAIL
   theo tiêu chí đã phê duyệt. Không xem test tổng hợp là nghiệm thu nghiệp
   vụ.

Không có thay đổi mã nguồn, mapping, dữ liệu nguồn hoặc mẫu trong lượt này.
Không chạy xuất 120 hộ và không commit.
