# Chuẩn hóa checkbox và trang in A4 — 2026-10-21

> Báo cáo lịch sử. Quy tắc kích thước/phân trang trong báo cáo này đã được
> thay bằng [6 người/trang ở 68%, giữ định dạng mẫu](SIX_MEMBERS_68_REVIEW.md).
> Quy tắc checkbox Unicode tiếp tục áp dụng.

## Nguyên nhân và cách sửa

Phiếu cũ có cùng chiều rộng 1108,5 pt nhưng chiều cao vùng hiện khoảng
1003,5 pt (7 khối) và 1638,75 pt (14 khối). `FitToPagesWide/Tall=1`
chọn tỷ lệ theo chiều cao nên dạng 14 khối nhỏ hơn rõ rệt. Lề kế thừa
không đối xứng, chưa căn giữa. Vùng in đã là A1:AG87; UsedRange kéo tới
hàng 94 nhưng không phải nguyên nhân chính; không có Shapes.

PDF cũ đo được độ rộng bảng/vùng in 83,2% và 50,3%, sai lệch tâm 47,17
và 181,05 pt. Chỉ đổi lề không giải quyết được tỷ lệ khung bảng.

Trên **bản xuất**, chuẩn hóa độ rộng cột bằng hệ số 1,7; dạng 7 khối dùng
dòng đối tượng cao 30 pt, dạng 14 khối dùng 15 pt. Hai dạng có cùng tổng
chiều cao hiện 1307 pt. Chỉ ẩn ba hàng phân cách trống 44:46 và các khối
8–14 khi hộ tối đa 7 người. Không xóa hàng, đổi vùng gộp hoặc dòng năm học.
Giữ nguyên cỡ chữ gốc toàn bảng; chỉ ShrinkToFit riêng ô có nội dung vượt
chiều rộng, kèm cảnh báo LONG_TEXT_FITTED. PDF bị thiếu nội dung trường
đã điền sẽ bị chặn, không nhận vào PDF tổng hợp.

Checkbox Nữ và hai vị trí cư trú dùng U+2612 ☒ / U+2610 ☐ với font
Segoe UI Symbol. Chỉ đánh dấu nữ khi nguồn xác định nữ; cư trú tối đa một
ô. Các trường khác như CNKT giữ nguyên cách ghi cũ. Kiểm tra cả giá trị
và font sau khi lưu Excel.

Quét mẫu gốc chỉ đọc tìm được đúng 16 ký hiệu q: 14 ô C ở dòng Nữ dùng
Wingdings và hai ký hiệu trong chuỗi Z2. Đây là mã ký tự phụ thuộc font,
có thể hiện thành q/ký hiệu khác khi thay nội dung hoặc mất font rich text;
bản xuất thay toàn bộ 16 vị trí bằng Unicode và font tường minh.

## PageSetup dùng chung Excel/PDF

- A4 (PaperSize=9), ngang (Orientation=2).
- PrintArea=$A$1:$AG$87; Zoom=False; FitToPagesWide=1; FitToPagesTall=1.
- Bốn lề 0,5 cm (14,1732 pt); header/footer 0; nội dung header/footer rỗng.
- Căn giữa ngang và dọc, reset ngắt trang thủ công trên bản xuất.
- PDF xác minh rồi áp dụng cùng PageSetup, không nhân độ rộng cột lần nữa.
- Hộ trên 14 người tiếp tục chia 14 người/sheet; mỗi sheet một trang.

## Bộ xuất giả lập được kiểm tra

Thư mục bàn giao: `output/testing/gui_batch_pdf_20261021_022346/`.
Chi tiết số đo nằm trong `VERIFICATION.json`; có 5 Excel, 5 PDF riêng,
PDF tổng hợp và 6 ảnh render trong `previews/` (đã xem đủ sáu trang).

| Hộ | Số sheet/trang PDF | Độ rộng bảng/vùng in | Chữ tên 10,5 pt đo trong PDF |
| --- | --- | --- | --- |
| 1 người | 1 / 1 | 99,94% | 4,56 pt ≈ 43,43% |
| 7 người | 1 / 1 | 99,94% | 4,56 pt ≈ 43,43% |
| 8 người | 1 / 1 | 99,94% | 4,56 pt ≈ 43,43% |
| 14 người | 1 / 1 | 99,94% | 4,56 pt ≈ 43,43% |
| 15 người | 2 / 2 | 99,94% cả hai | 4,56 pt ≈ 43,43% |

Tất cả trang A4 ngang, sai lệch tâm bảng 0,36 pt. Ký hiệu ☒/☐ trích xuất
đúng Unicode và hiển thị đúng trên ảnh render, không có q/ô đen/ký tự thay
thế. Giữ đủ năm dòng năm học, vùng ký, khối trống, đường bảng/vùng gộp.
Test COM/PDF dùng thêm bộ stress dài để đối chiếu các giá trị đã ánh xạ.
Bộ bàn giao điền cả lớp 10A1, trường THPT Kiểm thử, tốt nghiệp THCS/2024,
tôn giáo Không và các trường bổ sung để rà soát trực tiếp các vùng này.

Giới hạn thực tế: giữ 14 khối × 5 dòng trên một trang A4 vẫn làm chữ nhỏ;
đã loại bỏ khác biệt tỷ lệ giữa các trang và khoảng trống lớn bên phải,
không tuyên bố chữ đạt kích thước đọc của văn bản thông thường. Chuỗi dài
được vừa riêng ô có thể còn nhỏ hơn, có cảnh báo để rà soát.

Kiểm tra nguồn chỉ đọc: 427 người; không phát hiện ngày lịch không hợp lệ
trong các ngày đầy đủ, một người thiếu thành phần. Chuỗi số nguyên dạng
`2024.0` được hiểu đúng khi **kiểm tra**, không sửa chuỗi hoặc dữ liệu nguồn.
Ngày bất khả thi, số lẻ/không phải số có cảnh báo INVALID_BIRTH_DATE; test
ngày 31/04 xác minh cảnh báo và dữ liệu gốc không đổi.

PDF tổng hợp:
`output/testing/gui_batch_pdf_20261021_022346/pdf/dot_e8b3efbedf90413ab766dab7cc4f3782/TongHop.pdf`.

Không đổi nguồn, mẫu, mapping hoặc snapshot; không xuất hộ thật, không commit.

## Kiểm thử

Môi trường dev: `venv\Scripts\python.exe -m pip install -r requirements-dev.txt`.
Lệnh đầy đủ: `venv\Scripts\python.exe -m unittest discover -s tests -v`.
Kết quả: **50 PASS, 0 FAIL, 0 SKIP, 0 ERROR**, 576,026 giây.

File sửa trong lượt này: `src/print_layout.py`, `src/excel_exporter.py`,
`src/export_workflow.py`, `tests/test_print_layout.py`,
`tests/test_approved_mapping.py`, `tests/test_excel_exporter.py`,
`tests/test_export_workflow.py`, `scripts/verify_gui_batch_pdf.py`,
`requirements-dev.txt`, các tài liệu trạng thái/báo cáo.
