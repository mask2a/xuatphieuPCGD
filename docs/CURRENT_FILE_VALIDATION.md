# Xác minh workbook Excel hiện tại

Ngày kiểm tra: 2026-10-20
Phương pháp: Microsoft Excel COM, mở cả hai workbook với `ReadOnly=True`,
`UpdateLinks=0`; không lưu hoặc chỉnh sửa tệp.

## Kết quả đối chiếu cấu trúc

| Hạng mục | Kết quả hiện tại | So với audit trước |
|---|---|---|
| Nguồn — sheet | `MauNhapLieu`, `Sheet1`, `DuLieu`, `HuongDan`; `Sheet1` rỗng | Khớp danh sách trong `DATA_ANALYSIS.md` |
| Nguồn — UsedRange | `MauNhapLieu!A1:AY1059` | Khớp audit trước |
| Nguồn — bản ghi / hộ | 427 bản ghi dữ liệu từ hàng 5; 427 mã hộ không rỗng, 120 mã duy nhất | Khớp kết quả 427 nhân khẩu / 120 hộ |
| Nguồn — cột | Tiêu đề và các cột A:AX được đọc đúng; AY nằm trong UsedRange nhưng không thuộc cấu hình cột nguồn | Khớp danh mục cột trong `DATA_ANALYSIS.md` |
| Mẫu — sheet / UsedRange | `Phiếu điều tra 1 mặt`; `$A$1:$AG$94` | Khớp `TEMPLATE_MAPPING.md` và `SCHOOL_YEAR_MAPPING_AUDIT.md` |
| Khối thành viên | 14 vị trí, bắt đầu tại `9, 14, 19, 24, 29, 34, 39, 47, 52, 57, 62, 67, 72, 77` | Khớp cấu hình và audit |
| Nhãn năm học | Đúng 70 ô neo cột F; tất cả merge với G; giá trị mẫu từ `2025 - 2026` đến `2029 - 2030` | Khớp audit |
| Vùng gộp | 442 vùng | Khớp audit |
| Ô đích giáo dục đã duyệt | B5/R → H; B6/T → K; B7/V → P và X → R, ở 14 hàng bắt đầu nêu trên. Merge H:J và K:O đúng; P/R không gộp như cấu hình | Khớp các ô đích trong `template_mapping.json`; các ô kiểm tra mẫu hiện trống |
| Ô cố định đã duyệt | Xã A1 (`A1:G1`), mã hộ Z1 (`Z1:AG1`), chủ hộ H3 (`H3:Y3`), điện thoại Z3 (`Z3:AG3`) | Khớp cấu hình |
| Trang in | A4 ngang (`PaperSize=9`, `Orientation=2`), Zoom 68, Fit 1×1; PrintArea rỗng | Khớp audit; vùng in rỗng là trạng thái của mẫu gốc, exporter đặt vùng in trên bản xuất |
| Xác nhận | Bắt đầu hàng 82; các ô năm tại F82, J82, N82, R82, W82 lần lượt là năm 2025–2029; nội dung xác nhận gộp `AA82:AG83` | Khớp audit |
| External links | Không có | Không thấy khác biệt |

Cấu hình reader hiện khai báo các cột nguồn A:AX (50 cột); UsedRange tới AY
bao gồm một cột ngoài cấu hình này. Các cột được dùng trực tiếp bởi mapping
đã duyệt là B:F, G:I, K:L, N, Q, R, T, V, X, AU:AV và AW; toàn bộ cấu hình
cột cụ thể nằm trong `config/settings.json`. COM kiểm tra tiêu đề và số ô có
dữ liệu mà không đưa dữ liệu định danh cá nhân vào báo cáo.

## Snapshot phiên bản đã được người dùng xác nhận

| Tệp | Đường dẫn | Kích thước | Mtime UTC | SHA-256 chuẩn |
|---|---|---:|---|---|
| `dulieumau.xls` | `C:\Users\Admin\Desktop\VSCODE PROJECT\PHOCAPGD\dulieumau.xls` | 623104 bytes | `2026-10-19T14:25:01+00:00` | `e7ce0b359bc516d54463a29ab88142ff355b735bda6408f77566dc5e0d3b262a` |
| `maubaocao.xls` | `C:\Users\Admin\Desktop\VSCODE PROJECT\PHOCAPGD\maubaocao.xls` | 54272 bytes | `2026-10-18T12:52:27+00:00` | `963a75a4ae9258ff40f45006f97a9750c31023e0cee7bb4c115d938c52a9434e` |

Người dùng xác nhận sử dụng hai tệp hiện tại. Snapshot được lưu tại
`config/approved_workbook_snapshot.json`; test
`test_approved_workbook_snapshot_matches_current_files` kiểm tra đường dẫn,
kích thước, mtime UTC và SHA-256 riêng biệt với kiểm thử Excel COM chức năng.

Trước khi snapshot được phê duyệt, test
`SchoolYearLabelTests.test_updates_only_verified_labels_and_preserves_workbook_structure`
đã FAIL tại assertion hash mẫu. Test đó hiện chỉ kiểm thử hành vi cập nhật
nhãn năm qua Excel COM trên bản sao; không còn phụ thuộc digest phiên bản.
Kiểm tra version/snapshot tách biệt, nên thay đổi phiên bản được phát hiện
riêng mà không làm test logic nhãn thất bại.

## Ảnh hưởng và bước tiếp theo đề xuất

- Exporter kiểm tra sheet, số vùng gộp, định dạng, thiết lập trang và các ô
  mapping trên chính mẫu đang được chọn; các tiêu chí cấu trúc đo được hiện
  tại khớp cấu hình. GUI đọc nguồn vẫn thu được 427 bản ghi và 120 hộ.
- Hai digest trước đây (`dulieumau.xls`: `5382ace8...`; `maubaocao.xls`:
  `89354830...`) vẫn được lưu trong lịch sử tài liệu, không còn là snapshot
  hiện hành. Người dùng đã xác nhận snapshot mới trong
  `config/approved_workbook_snapshot.json`.
- Test snapshot so sánh cả metadata và digest với cấu hình đã duyệt; test
  COM tiếp tục kiểm tra 70 nhãn, toàn bộ ô khác không đổi, cấu trúc gộp,
  định dạng, chiều cao/chiều rộng, thiết lập in, vùng ký và external links.

## Kiểm thử liên quan

- Chạy riêng test snapshot và test COM chức năng: **2 PASS, 0 FAIL**.
- Chạy toàn bộ `unittest discover -v`: **36 PASS, 1 FAIL, 0 SKIP**
  (37 test). Test cập nhật 70 nhãn COM PASS. Lỗi duy nhất là
  `test_settings_leave_commune_unconfigured_and_load_configured_school_year`:
  test kỳ vọng `commune_name` rỗng nhưng cấu hình làm việc hiện tại là
  `"PDP"`. Không sửa `config/settings.json` vì đây là cấu hình địa phương
  đang có sẵn; lỗi không liên quan đến snapshot hoặc test nhãn năm.
- Kiểm tra JSON snapshot và cấu hình PASS; `git diff --check` PASS.
- Không sửa nguồn, mẫu, mapping hoặc dữ liệu thật; không xuất hộ thật.
