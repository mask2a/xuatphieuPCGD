# Đo và tối ưu tốc độ xuất — 2026-10-21

## Tùy chọn bỏ kiểm tra PDF mới

GUI có checkbox **Kiểm tra khi xuất PDF**, mặc định bỏ tích. Kiểm tra khi
tạo Excel không đổi. Bỏ tích không nghiệm thu lại template/Excel hoặc đối
chiếu bố cục/nội dung PDF; vẫn đọc được PDF, ghi nhận số trang thực tế và
gộp đúng thứ tự, không ghi đè. Tích bật lại đầy đủ các kiểm tra trước đây.

Đo cùng một hộ giả lập với `--skip-pdf-validation`: Excel 17,28 giây;
PDF **3,43 giây**; tổng 20,72 giây, so với chế độ đầy đủ trước đó khoảng
32,5–33,4 giây (PDF khoảng 15 giây). Đây là mức giảm kiểm tra theo lựa
chọn người dùng, không phải cùng mức kiểm tra; không suy ra tốc độ 38 hộ.
Log: `EXPORT_PDF_FAST_MODE.json`. Hồi quy phần tùy chọn: 10 test PASS,
0 FAIL; hai chế độ đều tiếp tục khi một hộ lỗi COM. Không chạy lại toàn bộ suite.

Đo tại public workflow `export_excel_batch` → `export_pdf_batch`, Excel COM
thật, một hộ giả lập, mẫu hiện tại. Dữ liệu thử nằm trong TemporaryDirectory
và tự xóa; không tạo đợt thử trong output chính thức.

## Bằng chứng trước/sau

| Chỉ số | Trước | Sau |
|---|---:|---:|
| Tạo và nghiệm thu Excel | 22,67 giây | 18,37 giây |
| Nghiệm thu và xuất PDF | 20,76 giây | 15,04 giây |
| Tổng thời gian | 43,43 giây | 33,42 giây |
| Số lần gọi COM Invoke | 33.969 | 25.689 |
| Thời gian trong `_com_cell_format` | 18,22 giây | 7,27 giây |

Thời gian giảm khoảng **23%**, số lần gọi COM giảm khoảng **24%**.
Baseline lần đầu độc lập là 42,12 giây, phù hợp lần lặp trước sửa.
Đây là phép đo một hộ trên máy hiện tại với profiler bật, không phải cam kết
cho mọi máy hay kết quả benchmark đủ 38 hộ/dữ liệu lớn.

Lần xác nhận thứ hai sau hồi quy: **32,50 giây**, vẫn **25.689** lần gọi
COM, PASS ngưỡng 29.000 và hash bảo vệ không đổi. Log:
`EXPORT_PERFORMANCE_AFTER_REPEAT.json`. Hai lần sau sửa giảm thời gian
khoảng 23–25% so với baseline 43,43 giây.

Các giả thuyết đã xếp hạng: (1) đọc định dạng qua COM lặp lại, (2) gán lại
PageSetup, (3) mở lại workbook/khởi động Excel. Profiler xác nhận (1) là
nút thắt lớn; chỉ thay đổi (1), không thay PageSetup hay số phiên Excel.

## Thay đổi

`src/excel_exporter.py::_com_cell_format` trước đây gọi lại
`cell.Borders.Item(index)` ba lần để lấy ba thuộc tính của cùng một cạnh.
Nay giữ đối tượng Borders/Border trong phạm vi **một ô**, vẫn đọc đủ
LineStyle, Weight, ColorIndex của cả bốn cạnh. Font, giá trị, hàng/cột,
vùng gộp và kiểm tra PDF không bị bỏ qua. Không cache kết quả định dạng
giữa các ô hoặc workbook, không dùng chung COM qua các luồng.

## Lặp lại phép đo

```powershell
.\venv\Scripts\python.exe scripts/benchmark_export.py --max-seconds 35 --max-com-calls 29000
```

Lệnh trước sửa với ngưỡng 35 giây đã FAIL (43,43 giây); sau sửa PASS
(33,42 giây, 25.689 lần gọi COM). Có thể chỉ dùng `--max-com-calls 29000`
để tránh kết luận sai khi máy đang bận. Không coi 35 giây là SLA của ứng dụng.
Log số đo chi tiết: `EXPORT_PERFORMANCE_BEFORE.json`, `EXPORT_PERFORMANCE_AFTER.json`.

Hash nguồn, mẫu và mapping trước/sau phép đo giữ nguyên. Không commit.
Khi dùng GUI cần khởi động lại để nạp mã mới.

Hồi quy sau tối ưu: **11 PASS, 0 FAIL**, 312,805 giây. Gồm cố ý sửa
đường viền (vẫn bị từ chối dù reseal hash), font lớp dài, số trang/vùng ký,
tiếp tục sau lỗi COM, tên đầu ra, chống ghi đè và GUI worker. Không chạy
lại toàn bộ suite; không đồng nhất số này với các test dữ liệu cũ đã biết lỗi.
