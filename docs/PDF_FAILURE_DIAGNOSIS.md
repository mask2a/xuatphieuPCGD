# Chẩn đoán lỗi PDF — 2026-10-21

## Cập nhật sau khi người dùng duyệt giảm font ô thiếu chỗ

- Chỉ ô **lớp học dài** trên bản xuất giảm xuống 8 pt khi không đủ chiều
  ngang ở font mẫu; giữ nguyên nội dung, cỡ ô, xuống dòng, hàng/cột và vùng
  gộp. Ô lớp ngắn/trống giữ nguyên font; không sửa workbook gốc.
- Nghiệm thu định dạng vẫn kiểm tra chính xác cỡ font dự kiến cho từng ô
  lớp theo giá trị và kích thước thực tế, không cho phép tùy ý đổi font.
- Test mới tại public workflow: RED khi ô còn 11 pt, GREEN sau sửa; Excel
  giữ nguyên chuỗi `11 CHUYÊN TRUNG`, PDF đủ chữ và đúng 1 trang.
- Đã xuất lại hộ 110 thực tế, xem bản render: chữ lớp đầy đủ ở 8 pt; các
  ô trống vẫn 11 pt. Giữ nguyên 37 PDF đã đạt kiểm tra, đối chiếu nội dung
  cả 38 PDF với nguồn hiện tại rồi gộp đúng thứ tự: **38 hộ, 38 trang, 0 lỗi**.
- Kết quả chính thức mới:
  `output/pdf/dot_fbf45f4d5f844c89b53308f426b0f77c/TongHop.pdf`.
  JSON đầu vào/hash và ảnh render cạnh file PDF. Hash nguồn/mẫu/mapping
  trước/sau không đổi. Không commit.
- Các ghi chép 37/38 và xung đột font bên dưới là trạng thái trước xác nhận
  mới của người dùng, được giữ để truy vết bằng chứng.
- Sau thay đổi font: chạy nhóm hồi quy PDF/COM, PageSetup, ngày sinh và
  validator nội dung: **10 PASS, 0 FAIL**, 157,689 giây. Chưa chạy lại toàn
  bộ suite sau thay đổi nhỏ này; số 59 test bên dưới là lượt trước thay font.

## Bằng chứng trước sửa

- Workbook thực tế: `output/excel/dot_3eeb97836ea44f8396e173ac88374ee5/ho_TNg_QTRI_T18_12.xlsx`.
- Chạy độc lập Excel COM `ExportAsFixedFormat`, không GUI, không gộp:
  `output/diagnostics/pdf_20261021_153440/DIAGNOSIS.json`.
- Một sheet hiện, PrintArea `$A$1:$AG$87`, A4 ngang, Zoom 68%, FitWide/Tall 1.
  Excel báo trước render 1 trang nhưng PDF thực tế có **2 trang**, đọc được,
  không mã hóa, 207598 byte. Trang 2 chỉ còn phần cuối vùng ký chủ hộ.
- Tổng chiều cao in 858,75 pt × 68% = 583,95 pt, vượt chiều cao in A4
  sau lề 0,5 cm (566,93 pt). Zoom số cố định khiến FitToPages không có tác dụng.
- Thử riêng MapPaperSize=False vẫn 2 trang: loại trừ giả thuyết đổi A4/Letter.
  Log: `output/diagnostics/pdf_20261021_153554/DIAGNOSIS.json`.
  GET.DOCUMENT(50) sau render là 2, trước render là 1: bộ đếm trước render bị cache.
- Validator đúng khi từ chối PDF 2 trang thay vì 1; traceback lưu ở
  `output/diagnostics/pdf_20261021_153554/VALIDATOR_TRACEBACK.txt`.
  Không phải lỗi COM xuất file, PDF hỏng hay PDF mã hóa.

## Sửa đúng nguyên nhân

- `src/print_layout.py::page_setup`: Zoom=False, FitToPagesWide/Tall=1,
  A4 ngang, lề bốn phía 0,5 cm, header/footer 0, căn giữa ngang/dọc.
- Mỗi sheet chứa tối đa **7 người**, vẫn giữ đủ 7 khối kể cả khối trống và
  vùng ký. Chỉ ẩn các khối dư của mẫu ở hàng 44:81, không đổi kích thước
  cột/hàng, font, viền hay vùng gộp. Quy tắc mới thay 6 người/68% trước đây.
- `src/export_workflow.py::export_pdf_batch`: kiểm tra bộ đếm Excel sau
  render thay vì dựa vào bộ đếm cache trước render. Vẫn bắt buộc mỗi sheet
  đúng 1 trang, PDF đúng tổng số sheet, A4 ngang, đọc được, không mã hóa,
  không mất nội dung; không bỏ bất kỳ điều kiện số trang nào.
- Kiến trúc A: sao chép nhiều sheet, mỗi sheet 7 thành viên; workbook có
  thể 1, 2, 3... trang. Giữ nguyên thứ tự thành viên và thứ tự hộ khi gộp.

## TDD và kiểm chứng

`tests/test_pdf_regression.py` kiểm thử public workflow Excel → PDF bằng COM thật.
Đã thấy RED trước sửa (1 người tràn 2 trang), GREEN sau sửa; ca 7 người
cũng RED khi còn chia 6, GREEN sau chuyển 7. Có test tiếp tục xuất hộ sau
khi một lần COM thất bại và test PDF sai số trang vẫn bị từ chối.

Hộ thực tế TNg_QTRI_T18_12 sau xuất lại: **1 trang**, đủ vùng ký:
`output/pdf_recovery/20261021_154419_one/RECOVERY.json`.
Các ca 1/7/8/14/15 người: **1/1/2/2/3 trang**, 5 thành công, 0 lỗi;
đã xem render cả 9 trang, không mất vùng ký:
`output/pdf_recovery/20261021_154546_boundaries/RECOVERY.json`.

Các file nguồn/mẫu thay đổi hash giữa lượt trên và lượt chạy tiếp nên gate
đã chặn chạy 38 hộ. Không tự cập nhật snapshot; chạy lại ca biên trên
phiên bản hiện tại trước khi chạy hàng loạt. Kết quả cuối được bổ sung sau
khi hoàn tất, không dùng kết quả cũ để khẳng định phiên bản mới đã PASS.

## Lỗi nội dung phát hiện ở lượt 38 hộ

Lượt `output/pdf_recovery/20261021_155518_batch/RECOVERY.json`: Excel
38/38 thành công; PDF 3 thành công, 35 bị từ chối. Không còn lỗi số trang.

Đã xuất trực tiếp hộ TNg_QTRI_T18_14: PDF 1 trang, không mã hóa, 215006
byte; log ở `output/diagnostics/pdf_20261021_160401/DIAGNOSIS.json`.
Ô R9 chứa nguyên chuỗi `2020-2021`, Excel xuống dòng thành `2020-` / `2021`.
Trích xuất PDF kiểu layout chèn văn bản ô bên cạnh vào giữa hai dòng;
trích xuất kiểu plain giữ nguyên `2020-2021`. Đây là lỗi validator nội dung,
không phải mất dữ liệu. Traceback trước sửa: `CONTENT_TRACEBACK.txt` cùng thư mục.

`validate_pdf_content` nay đối chiếu sự có mặt của giá trị bằng cả hai cách
trích xuất trên **cùng trang**; thứ tự thành viên vẫn chỉ theo layout để tránh
thứ tự stream PDF nhóm theo font. Không nối dữ liệu giữa các trang, không
bỏ trường cần kiểm tra. Test xuống dòng RED trước sửa, GREEN sau sửa;
test mất chữ thật/sai thứ tự/sai số trang vẫn PASS.

Có trường hợp cắt chữ thật cần giữ chặn: hộ `TNg_QTRI_T18_ 110`, thành viên
thứ 3, lớp `11 CHUYÊN TRUNG`. PDF chỉ hiện `11` / `CHUYÊ` vì ô mẫu không đủ
chỗ. Log và bản render: `output/diagnostics/pdf_20261021_160502/`.
Không đổi font, độ rộng cột, chiều cao dòng hoặc vùng gộp để che lỗi này.
Đây là xung đột giữa yêu cầu giữ nguyên mẫu và hiển thị đủ chuỗi dài;
cần người dùng duyệt cách bố trí trước khi xử lý định dạng.

Lượt kế tiếp `output/pdf_recovery/20261021_160930_batch/RECOVERY.json`:
35/38 PDF thành công; hai hộ 06/07 còn bị từ chối do lớp là **số** 6.0 trong
nguồn, ô Excel General/PDF hiển thị đúng `6`. Log riêng tại
`output/diagnostics/pdf_20261021_161355/` và `pdf_20261021_161357/`.
Sau test RED, validator chuẩn hóa riêng số float nguyên khi so sánh bản in;
không chuẩn hóa chuỗi `"6.0"`, không đổi dữ liệu hay cách exporter ghi ô.
Test GREEN và hai PDF thực tế PASS; hộ 110 vẫn bị từ chối. Các test mất
chữ thật, đảo thứ tự và sai số trang tiếp tục PASS. Lượt toàn bộ test đang
chạy đã được dừng trước khi sửa này; không tính lượt dở dang vào kết quả.

## Kết quả chạy lại cuối cùng

- Ca biên mã cuối: `output/pdf_recovery/20261021_161606_boundaries/RECOVERY.json`:
  1/7/8/14/15 người → 1/1/2/2/3 trang, 5 Excel và 5 PDF PASS, 0 lỗi.
- Chạy lại đủ 38 hộ: `output/pdf_recovery/20261021_161747_batch/RECOVERY.json`.
  Excel 38 thành công/0 lỗi; PDF **37 thành công/1 lỗi**. Hộ lỗi duy nhất là
  `TNg_QTRI_T18_ 110`, thành viên 3, lớp `11 CHUYÊN TRUNG` bị cắt thật.
- PDF tổng hợp **37 trang**, không chứa hộ 110 bị từ chối:
  `output/pdf_recovery/20261021_161747_batch/pdf/dot_e1d9f23270894e80a8b72346874f7a1e/TongHop.pdf`.
  Toàn bộ trang A4 ngang, không mã hóa, nội dung từng trang khớp PDF riêng;
  thứ tự đúng danh sách hộ đã thành công. Đã render và xem 37 trang;
  số đo/các ảnh lưu ở thư mục `previews/` cạnh báo cáo đợt.
- Hộ đầu TNg_QTRI_T18_12: trước sửa 2 trang/thay vì 1 dự kiến; sau sửa 1 trang.
- Hash nguồn, mẫu và mapping trước/sau lượt cuối giống nhau. Không sửa
  workbook gốc hoặc mapping, không commit, không bỏ kiểm tra mất chữ.
- Chưa tuyên bố đạt 38/38. Trường hợp hộ 110 cần phương án hiển thị chuỗi
  dài được người dùng chấp thuận; không tự đổi định dạng trong nhiệm vụ này.

Các file mã chính sửa: `src/print_layout.py` (7 người và Fit),
`src/excel_exporter.py` (nghiệm thu hàng ẩn phù hợp 7 khối),
`src/export_workflow.py` (đếm trang sau render, xử lý trích xuất xuống dòng,
so sánh số nguyên). Bổ sung hồi quy `tests/test_pdf_regression.py` và
`tests/test_export_workflow.py`; cập nhật kỳ vọng phân trang trong test hiện có.
Script chẩn đoán: `scripts/diagnose_pdf_export.py`; công cụ kiểm chứng theo
từng giai đoạn: `scripts/verify_pdf_recovery.py`.

## Toàn bộ test mã cuối

Lệnh tương đương `python -m unittest discover -s tests -v` được chạy qua
`output/diagnostics/run_full_suite.py` để lưu log và hash bảo vệ.
Log gốc: `output/diagnostics/suite_20261021_162159/unittest.log`;
JSON có đầy đủ traceback: `output/diagnostics/suite_20261021_162159/RESULT.json`.

**59 test methods: 54 PASS, 5 FAIL, 0 ERROR, 0 SKIP**; thời gian 1100,861 giây.
Unittest ghi `failures=6` vì một method có hai subtest thất bại, không phải
6 method khác nhau. Báo cáo JSON đếm method riêng và giữ nguyên cả 6 traceback.

- Bốn method thuộc reader, analysis_report, household_grouper và GUI còn
  cố định bộ dữ liệu 427 người/120 hộ; nguồn hiện tại 141 người/38 hộ.
  Không đổi dữ liệu hoặc thay số kỳ vọng chỉ để làm test xanh.
- Method `test_updates_only_verified_labels_and_preserves_workbook_structure`
  có hai subtest 2026-2027/2027-2028 còn kỳ vọng PrintArea rỗng ở dòng 269;
  bản sao của mẫu hiện tại giữ `$A$1:$AG$94`. Không sửa mẫu để thỏa test cũ.
- Tất cả test hồi quy PDF, bộ ca COM nhiều sheet gồm 21 người, kiểm tra
  thứ tự/gộp/chống sửa file/chống ghi đè và tiếp tục sau lỗi COM đều PASS.
- Hash nguồn, mẫu và mapping trước/sau toàn bộ test giống nhau; Excel COM
  đã được đóng sau khi hoàn tất. Không công bố bộ test hoàn toàn xanh.
