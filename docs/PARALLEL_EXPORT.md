# Xuất nhiều tệp song song — 2026-10-21

GUI có mục **Xuất song song (1–4)**, mặc định 2, được lưu khi bấm Lưu cấu hình.
Mỗi tiến trình sở hữu một Excel COM riêng; không chia sẻ COM giữa các luồng.
Tác vụ nền của GUI vẫn cập nhật tiến độ, nhật ký và thống kê.

Với nhiều tệp nguồn, xử lý tối đa số nguồn đã chọn cùng lúc. Chọn 1 để chạy
tuần tự; một tệp nguồn vẫn chạy tuần tự các hộ, không tạo tiến trình dư.
Hàng đợi chỉ nạp thêm nguồn khi có tiến trình rảnh. Tên thư mục được giữ chỗ
trước khi chạy để tránh trùng tên; kết quả trả về theo thứ tự nguồn, PDF gộp
theo thứ tự hộ. Lỗi thông thường của một nguồn được báo riêng và không làm
dừng các nguồn tiếp theo. Mỗi worker đóng workbook và Excel trong finally.

**Kiểm tra khi xuất PDF** được truyền cho từng worker. Bỏ tích để chuyển
Excel đã xuất trực tiếp sang PDF rồi gộp, không nghiệm thu lại nội dung/định dạng.
Vẫn phải đọc PDF để gộp an toàn; không ghi đè file cũ.

## Đo trên cùng hai nguồn giả lập, mỗi nguồn một hộ

| Chế độ | Excel | PDF | Tổng |
|---|---:|---:|---:|
| 1 tiến trình, bỏ kiểm tra PDF | 29,88 s | 6,86 s | 36,74 s |
| 2 tiến trình, bỏ kiểm tra PDF | 17,87 s | 4,54 s | 22,41 s |

Giảm khoảng 39% thời gian tổng, tương đương nhanh hơn 1,64 lần trên phép đo
này. Không suy diễn cùng mức tăng cho mọi tệp hoặc cấu hình máy. Mặc định 2
để hạn chế bộ nhớ/CPU; mức 3–4 cần người dùng chọn tùy tài nguyên.
Log: `PARALLEL_EXPORT_BEFORE.json`, `PARALLEL_EXPORT_AFTER.json`.
Profiler chỉ ghi lời gọi COM ở tiến trình cha, không thể dùng để so số lời gọi
COM tổng giữa hai chế độ. Hash nguồn, mẫu và mapping trước/sau giữ nguyên.

```powershell
.\venv\Scripts\python.exe scripts/benchmark_export.py --sources 2 --workers 2 --skip-pdf-validation
```

Kiểm thử dùng thư mục tạm, không xuất dữ liệu thật. Các kiểm tra bao phủ
tiến trình thực sự chạy chồng thời gian, trùng tên nguồn/mã hộ, nguồn lỗi,
PDF nhiều trang, tiến độ đơn điệu và truyền/lưu tùy chọn từ GUI.

Kết quả hồi quy: **10 test riêng biệt PASS, 0 FAIL, 0 SKIP** sau khi chạy đúng
tên lớp test (lệnh đầu có một lỗi nạp test do gõ sai tên lớp; đã chạy lại đúng).
Gồm 2 test `test_parallel_export`, 2 test `test_pdf_option`, 5 test GUI liên quan
và `PdfPaginationRegressionTests.test_com_export_failure_does_not_poison_next_household`
(ca lỗi COM kiểm tra cả hai chế độ PDF). Chế độ đầy đủ trong worker vẫn từ chối
Excel bị sửa sau khi nghiệm thu; bỏ tích xuất được trực tiếp bản Excel hiện tại.
Không chạy lại toàn bộ suite trong lượt này. Kiểm tra cú pháp và `git diff --check`
PASS. Sau khi kết thúc không còn tiến trình EXCEL được liệt kê bằng Get-Process.
