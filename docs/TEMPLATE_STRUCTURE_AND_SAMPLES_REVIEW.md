# Kiểm tra cấu trúc mẫu và 6 workbook — 2026-10-21

## Chính sách hiện hành

Người dùng xác nhận đã mở/lưu lại mẫu và yêu cầu cấu trúc thực tế là điều
kiện chặn xuất. SHA-256, mtime và kích thước snapshot cũ không còn là điều
kiện nghiệm thu mẫu. Snapshot cũ được giữ nguyên để tham khảo lịch sử.
SHA-256 được ghi qua logger và trong các báo cáo JSON; không so hash để chặn.
Không đọc được hash chỉ ghi cảnh báo và tiếp tục xác minh cấu trúc.

Các kiểm tra cấu trúc dùng chung trong exporter và test:

- Sheet theo mapping, nhãn/merge của các mục thông tin hộ.
- Cả 14 khối, số thứ tự, các ô cơ bản, 70 ô năm F:G và năm cố định 2025–2030.
- Các đích đã duyệt và đích mới đúng merge; ô mới không có dữ liệu/công thức
  sẵn; checkbox Nữ đúng vị trí, không có công thức.
- Vùng ký AA82:AG83, các chuỗi năm ký 2025–2029, A4 ngang.
- 442 vùng gộp, không có external links. Giữ việc chụp và đối chiếu định
  dạng, kích thước hàng/cột, lề và vùng ký giữa mẫu và bản xuất.
- Giữ kiểm tra chặn ghi đè công thức và nghiệm thu workbook sau lưu.
- Test làm sai merge vẫn phải bị từ chối; test cấu trúc không còn phụ
  thuộc hash/path/mtime/kích thước snapshot cũ.

## Đối chiếu 6 file đã xuất trước

Thư mục: `output/testing/approved_mapping_20261021_010939/`.
Đọc bằng openpyxl và Excel COM `ReadOnly=True`, không lưu thay đổi.
Tái dựng đúng dữ liệu tổng hợp bằng `full_household` đã tạo bộ mẫu trước.
Đối chiếu kỳ vọng từng ô với dữ liệu tổng hợp và mẫu thực tế đã xác minh;
không dùng dữ liệu định danh thật.

- Tổng **20.097 ô**, 7 sheet, 50 thành viên tổng hợp.
- Thông tin hộ/Xóm/cư trú; họ tên/ngày sinh/quan hệ/Nữ/dân tộc/tôn giáo;
  5 nhãn năm; TN nghề; học xong/bỏ học/xóa mù chữ; loại khuyết tật,
  CNKT/KNHT; chi tiết hoàn cảnh và ghi chú đều đúng kỳ vọng.
- Lớp/trường/bậc và năm tốt nghiệp trong bộ 6 mẫu này là trống, khớp đầu vào.
  Các ca dữ liệu không trống vẫn được kiểm tra trong test exporter hiện có.
  Không coi kiểm tra ô trống là bằng chứng cho mọi giá trị nghiệp vụ có thể có.
- Các khối không dùng còn nguyên mẫu, không thêm dữ liệu; không xóa dòng.
  Bộ 7 khối đầu hiện ở phiếu 1–7 người; phiếu 8–14 hiện đủ 14 khối.
- Vùng gộp, chiều cao dòng/rộng cột, font/viền/fill/căn lề/protection và vùng
  ký khớp mẫu, ngoại trừ thay đổi font X có chủ đích và lỗi định dạng Z2 dưới đây.
- Font Arial (kèm family/charset phù hợp) ở dấu X Nữ thay Wingdings là
  thay đổi đã duyệt để hiển thị chữ X, không phải mất định dạng.
- Excel GET.DOCUMENT(50): hộ 1/5/7/8/14 đều 1 trang; hộ 15 có 2 sheet x 1 trang.

Phát hiện thực tế: **Z2 của file hộ 5, 7 và 14 người có NumberFormat `@`
thay vì `General`**. Không sai giá trị. Báo cáo đầy đủ trong
`EXISTING_SAMPLES_AUDIT.json`, ghi đúng 3 lệch định dạng này.

## Sửa lỗi và bộ kiểm chứng mới

Exporter nay lưu NumberFormat của Z2 trước khi thay ký tự và khôi phục sau
khi đánh dấu X. Bổ sung assertion định dạng Z2 trong test tích hợp.
Giữ nguyên 6 file cũ để đối chiếu; tạo bộ dữ liệu tổng hợp mới tại:

`output/testing/approved_mapping_20261021_012426/`

Gồm `ho_SYNTHETIC-01.xlsx`, `ho_SYNTHETIC-05.xlsx`, `ho_SYNTHETIC-07.xlsx`,
`ho_SYNTHETIC-08.xlsx`, `ho_SYNTHETIC-14.xlsx`, `ho_SYNTHETIC-15.xlsx`.

Đối chiếu lại **20.097 ô, 0 sai lệch**, kể cả NumberFormat Z2.
Phân trang vẫn 1/1/1/1/1/(1+1). Hash nguồn, mẫu, snapshot và các file được
kiểm tra trước/sau không đổi. Kết quả: `REVIEWED_SAMPLES_AUDIT.json`.

## Phạm vi thay đổi

`src/excel_exporter.py`, `src/approved_mapping.py`,
`tests/test_school_year_labels.py`, `tests/test_approved_mapping.py`;
thêm công cụ chỉ đọc `scripts/audit_existing_mapping_samples.py` và báo cáo.
Không đổi mapping, GUI, reader, grouper hoặc settings trong lượt này.
Không sửa workbook gốc, snapshot; không xuất hộ thật, không commit.
Không mở khóa Xuất tất cả/PDF, không tự tuyên bố nghiệm thu.

## Kiểm thử cuối

`.\venv\Scripts\python.exe -m unittest discover -v`:
**43 PASS, 0 FAIL, 0 SKIP, 0 ERROR**, thời gian 316,940 giây.
`py_compile` và `git diff --check` đều PASS. Không còn test chặn theo hash
snapshot cũ; các kiểm tra cấu trúc, dữ liệu, công thức và bảo toàn bản gốc
vẫn hoạt động. Dừng để người dùng kiểm tra kết quả.
