# Xuất Excel hàng loạt và PDF trên GUI — 2026-10-21

## Nguyên nhân trước đây

Hai nút bị đặt `disabled`, không có `command` hoặc handler. Exporter đã có
khả năng xử lý nhiều hộ nhưng GUI chỉ gọi adapter xuất thử một hộ. Chưa có
module chuyển/gộp PDF hoặc phụ thuộc đọc PDF.

## Cách sử dụng

1. Chạy `.\venv\Scripts\python.exe main.py`.
2. Chọn nguồn, mẫu, thư mục đầu ra, nhập xã/phường và đọc dữ liệu.
3. Bấm **Xuất tất cả Excel**. Không cần chọn từng dòng hộ; thao tác dùng
   toàn bộ dữ liệu đã nạp, không chỉ các dòng đang hiện sau tìm kiếm.
4. Sau khi Excel hoàn tất, bấm **Xuất PDF và tổng hợp**. Chỉ các Excel đã
   xuất và nghiệm thu thành công trong đợt gần nhất của phiên GUI được dùng.
   Có thể chạy lại bước PDF từ đợt này mà không xuất lại Excel.

Nút Excel mở khi có dữ liệu, PDF mở sau đợt Excel có kết quả hợp lệ. Cả hai
bị khóa khi có tác vụ nền. Đổi xã/phường cần đọc lại dữ liệu; đọc lại nguồn
xóa liên kết tới đợt Excel trước để tránh ghép dữ liệu khác phiên.

Mỗi lần xuất tạo thư mục riêng:

```text
<thư mục đã chọn>/
  excel/dot_<id>/ho_<mã>.xlsx + EXPORT_REPORT.json
  pdf/dot_<id>/ho_<mã>.pdf + TongHop.pdf + PDF_REPORT.json
```

Không ghi đè file của đợt cũ. Mã trùng sau chuẩn hóa của exporter bị báo
lỗi, không tự chọn một hộ hoặc ghi đè dữ liệu hộ khác. Mã thiếu dùng đúng
mã dự phòng của exporter. Các va chạm tên sau làm sạch được exporter từ
chối an toàn, những hộ khác vẫn tiếp tục.

## Luồng và bảo vệ

- Tái sử dụng `export_households`, mapping và nghiệm thu Excel hiện có.
  Cấu hình thao tác được chụp trong tệp tạm; không sửa cấu hình thật để xuất.
- Một phiên COM cho mỗi đợt Excel; một phiên COM cho mỗi đợt PDF. Mọi COM
  chạy trên worker sở hữu phiên, không trên thread Tkinter.
- PDF nghiệm thu lại **toàn bộ** đợt Excel trước khi bắt đầu chuyển; so
  fingerprint của bản Excel đã nghiệm thu để chặn file bị sửa xen giữa.
  Đây là bảo vệ đầu ra đã nghiệm thu, không phải khôi phục kiểm tra hash mẫu.
- Mẫu vẫn kiểm tra cấu trúc thực tế, hash chỉ log theo chính sách đã duyệt.
- Excel COM `ExportAsFixedFormat`, tôn trọng PrintArea. Mỗi sheet phải in
  đúng 1 trang; hộ >14 giữ sheet bổ sung. Không lưu lại workbook đầu vào.
- pypdf đọc lại mỗi PDF: số trang, trang có nội dung, A4 ngang. Chỉ PDF
  hợp lệ được gộp; giữ thứ tự mã hộ tăng dần theo phần số, không thay mã.
- Gộp qua tệp tạm, kiểm tra lại tổng trang rồi đổi tên; không ghi đè
  tổng hợp đã tồn tại. Lỗi từng hộ được ghi nhận; lỗi gộp báo riêng.
- Khôi phục `Application.MapPaperSize` và đóng workbook/Excel COM trong
  `finally`, dọn tệp tạm. Excel trên máy này tự đổi A4 sang Letter nếu để
  bật MapPaperSize; luồng PDF tắt tùy chọn trong phiên riêng rồi khôi phục.
  Tham chiếu: [Microsoft MapPaperSize](https://learn.microsoft.com/en-us/office/vba/api/excel.application.mappapersize).
- Tiến độ từng hộ, nhật ký, thống kê Excel/PDF thành công và lỗi, tổng
  trang tổng hợp nằm trên GUI; hai báo cáo JSON giữ kết quả sau khi đóng app.

## Bộ thử giữ lại

`output/testing/gui_batch_pdf_20261021_014204/`:

- 5 hộ tổng hợp DEMO-01, DEMO-07, DEMO-08, DEMO-14, DEMO-15.
- **5 Excel thành công, 5 PDF thành công, 0 lỗi; tổng hợp 6 trang.**
- Mỗi hộ <=14: 1 trang; hộ 15: 2 trang theo 2 sheet.
- Excel: `excel/dot_64f39eaa384d47bfbbd86b2ba6d9eeda/`.
- PDF: `pdf/dot_e09ed8faa9954e09ab9da44f789da5de/`.
- Render đủ 6 trang trong `previews/`, đã xem trực quan: khung, chữ, ô X,
  vùng ký còn nguyên, không có trang trắng; giá trị mẫu ngắn vừa ô.
- Hash nguồn/mẫu/mapping trước và sau không đổi (`VERIFICATION.json`).
- Bộ stress trước đó ở `gui_batch_pdf_20261021_013917` dùng chuỗi TEST dài
  và nhiều loại khuyết tật cùng lúc, không dùng làm bản xem cuối.

## Kiểm thử mới

- GUI bấm nút, worker riêng, chặn chạy đồng thời, xử lý lỗi giải phóng busy,
  cập nhật thống kê và chặn xuất khi xã/phường đã thay đổi.
- Excel/PDF COM thật: 1/7/8/14/15 người, thứ tự gộp, 6 trang, giữ bản gốc.
- Từ chối ghi đè PDF tổng hợp; không để sót PDF tạm.
- Excel bị sửa sau nghiệm thu: loại riêng hộ đó, tiếp tục hộ khác và gộp
  đúng 5 trang còn lại trong ca thử.
- Mã trùng bị báo lỗi, không dùng nhầm dữ liệu; mã dự phòng vẫn được nhận.

## Giới hạn còn lại

**Cập nhật sau rà soát A4/checkbox:** xem `PRINT_A4_CHECKBOX_REVIEW.md`.
Mô tả dưới đây ghi nhận bản thử 01:42, trước lần chuẩn hóa trang in;
bản xuất mới có hình học thống nhất và xử lý riêng ô nội dung dài.

- PDF trong GUI hiện dùng đợt Excel gần nhất **của phiên đang mở**, chưa
  có màn hình nhập lại một thư mục Excel bất kỳ sau khi khởi động lại.
- Bố cục và kích thước ô mẫu được giữ nguyên. Chuỗi dài ở cột hẹp có thể
  bị cắt theo chính cách Excel hiển thị (đã thấy ở bộ stress); không tự
  nới dòng/co chữ hoặc đổi mapping trong nhiệm vụ này. Phiếu 14 người có
  chữ nhỏ do phải giữ đúng một trang theo bố cục đã duyệt.
- Chỉ thử dữ liệu tổng hợp; chưa chạy xuất 120 hộ thật, không tự tuyên bố
  nghiệm thu nghiệp vụ.

Mã thay đổi: `gui.py`, `src/excel_exporter.py` (thêm callback tiến độ),
`src/export_workflow.py`, `requirements.txt` (pypdf), test và script kiểm chứng.
PyMuPDF chỉ dùng render kiểm thử, không cần cho chức năng GUI. Không đổi
mapping, reader, grouper, nguồn hoặc mẫu; không commit.

## Kết quả cuối

`.\venv\Scripts\python.exe -m unittest discover -v`: **48 PASS, 0 FAIL,
0 SKIP, 0 ERROR**, 431,080 giây. Ca báo cáo JSON mã hộ dự phòng được chạy
riêng thêm sau chỉnh metadata và PASS. `py_compile`, JSON báo cáo và
`git diff --check` PASS. Cờ kiểm chứng tính năng đã bật sau khi các ca COM
Excel/PDF và GUI tương ứng PASS; nút vẫn phụ thuộc dữ liệu/Excel hợp lệ.

Không ghi nhận lỗi vận hành trong các ca đã chạy. Các giới hạn về phiên
GUI và chuỗi dài trong ô mẫu vẫn như mục Giới hạn còn lại, không tự thay
đổi bố cục đã được duyệt.
