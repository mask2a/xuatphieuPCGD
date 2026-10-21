# Mốc sử dụng ổn định — 2026-10-21

Người dùng xác nhận bản hiện tại chạy ổn, tốc độ vẫn còn cần cải thiện.
Commit mốc này trước khi bắt đầu phiên bản tiếp theo; không triển khai
thay đổi chức năng mới trong lần lưu mốc.

- GUI đọc một file hoặc thư mục, reader có số luồng giới hạn.
- Xuất Excel/PDF qua Excel COM, mapping đã duyệt, tối đa 7 người/trang.
- Font ô lớp dài giảm cục bộ để đủ chữ, giữ cấu trúc mẫu.
- Đầu ra chia theo tên nguồn; PDF tổng hợp nằm ở thư mục output gốc.
  Xuất lặp thêm hậu tố, không ghi đè kết quả cũ.
- Có tiến độ, nhật ký, worker nền và chặn tác vụ xuất đồng thời.
- Có tối ưu tái sử dụng đối tượng đường viền COM trong phạm vi một ô;
  không bỏ kiểm tra nội dung/định dạng hoặc tăng phiên COM song song.

Kiểm chứng gần nhất: nhóm hồi quy 11 PASS, 0 FAIL; chưa chạy lại toàn bộ
suite. Lượt toàn bộ trước đó còn các test kỳ vọng dữ liệu/PrintArea cũ,
không khẳng định mọi test đều xanh. Chi tiết trong PROJECT_STATUS.md và
docs/EXPORT_PERFORMANCE.md.

Dữ liệu nguồn, mẫu Excel, PDF tham chiếu, venv và output là tệp cục bộ,
không đưa vào Git. Báo cáo JSON chứa bản ghi trích xuất nguồn cũng giữ
cục bộ. Cấu hình/mapping, mã nguồn, test và tài liệu bàn giao được lưu.
Không push lên remote trong yêu cầu này.
