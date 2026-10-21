# Đọc tệp/thư mục và xuất chính thức

Chạy GUI: `.\venv\Scripts\python.exe main.py`.

1. Chọn **Chọn tệp** để đọc một `.xls`/`.xlsx`, hoặc **Chọn thư mục** để
   đọc các tệp Excel trong thư mục và các thư mục con.
2. Chọn mẫu, thư mục kết quả, nhập xã/phường. Mục **Luồng đọc (1–8)**
   mặc định 4; chọn 1 nếu cần hạn chế dùng bộ nhớ. Lưu cấu hình nếu muốn
   dùng lại lần sau. Nhấn **Đọc dữ liệu**.
   Thư mục kết quả cần tách riêng, không trùng hoặc chứa tệp/thư mục nguồn.
3. Chọn một hộ rồi bấm **Xuất hộ đã chọn**, hoặc bấm **Xuất tất cả Excel**
   để xuất toàn bộ hộ đã nạp, kể cả các hộ không hiện do bộ lọc tìm kiếm.
4. Bấm **Xuất PDF và tổng hợp** để chuyển đợt Excel gần nhất trong phiên
   GUI đang mở. Đợt này có thể là một hộ hoặc toàn bộ hộ.

Thư mục kết quả chính thức:

```text
<thư mục đã chọn>/
  excel/to1/ho_<mã hộ>.xlsx
  excel/to1/EXPORT_REPORT.json
  pdf/to1/ho_<mã hộ>.pdf
  pdf/to1/PDF_REPORT.json
  to1.pdf
  excel/to2/ho_<mã hộ>.xlsx
  pdf/to2/ho_<mã hộ>.pdf
  to2.pdf
```

Luồng GUI không tạo `test_output` hoặc tự chuyển đầu ra vào `output/testing`.
Các script trong `scripts/` vẫn là công cụ kiểm chứng riêng, không phải
đường xuất của ứng dụng. Các tệp cũ được giữ nguyên, không tự xóa/di chuyển.

Tên nhóm lấy từ tên tệp nguồn bỏ phần mở rộng. Nếu chọn thư mục có nhiều
tệp, mỗi tệp có nhóm Excel/PDF và PDF tổng hợp riêng như ví dụ to1/to2.
Nếu thư mục chỉ nạp một tệp, dùng tên thư mục làm tên nhóm. Xuất hộ đã
chọn cũng theo quy tắc nguồn này. Tên trùng hoặc xuất lại được thêm `_2`,
`_3`...; không ghi đè kết quả cũ. PDF tổng hợp nằm trực tiếp tại thư mục
đầu ra do người dùng chọn. Không gộp chung các tệp nguồn chỉ vì mã hộ trùng;
mã hộ trùng trong cùng nguồn vẫn bị chặn.

## Xử lý nhiều tệp

- Reader đọc các workbook độc lập bằng pool 1–8 luồng, chỉ giữ số tác vụ
  đang chờ bằng số luồng. Một workbook vẫn do một luồng đọc; không chia
  các hàng của cùng workbook cho nhiều luồng.
- Tiến độ tách rõ bước đọc tệp, chuẩn hóa và nhóm hộ. GUI xử lý nhật ký
  theo lượt có giới hạn để không giữ vòng giao diện quá lâu khi có nhiều tệp.
- Kết quả được ghép theo thứ tự nguồn xác định, không theo thứ tự hoàn
  thành tác vụ. Nguồn trùng đường dẫn chỉ đọc một lần; tệp hỏng được báo
  riêng, không làm mất các tệp đọc thành công.
- Không quét vào thư mục kết quả, các thư mục hệ thống/venv đã loại trừ,
  liên kết thư mục, mẫu đang chọn hoặc tệp tạm. Khi bỏ cả nhánh thư mục,
  nhật ký ghi đường dẫn nhánh đó thay vì liệt kê mọi tệp bên trong.
- Mã hộ trùng giữa nguồn vẫn được báo nhưng xuất vào nhóm nguồn riêng,
  không âm thầm gộp hay ghi đè. Không đổi mã hộ. Mã trùng trong cùng
  nguồn vẫn bị chặn như trước.
- Nếu đổi đường dẫn nguồn hoặc xã/phường sau khi đọc, cần đọc lại trước
  khi xuất. Không cho chạy đồng thời hai tác vụ từ GUI.
- Excel COM vẫn chạy tuần tự trong worker riêng; không chia sẻ COM object
  giữa các luồng đọc. Đối chiếu kết quả Excel theo mã hộ đã chuyển từ vòng
  tìm kiếm lồng nhau sang tra cứu dictionary để tránh chi phí bậc hai.

Giữ mapping đã duyệt, tối đa 7 người/sheet/trang, A4 ngang Fit 1×1,
Zoom=False, lề 0,5 cm, font Dân tộc 8 pt,
checkbox Unicode và các kiểm tra nghiệm thu Excel/PDF. Không sửa dữ liệu
nguồn hoặc mẫu. GUI chưa hỗ trợ nhập lại một thư mục Excel cũ để chuyển
PDF sau khi khởi động lại.

## Trạng thái bàn giao

Quy tắc 7 người và Fit 1×1 thay thế 6 người/68% theo yêu cầu chẩn đoán
PDF mới nhất. Xem [bằng chứng và kết quả kiểm chứng](PDF_FAILURE_DIAGNOSIS.md).
Sau khi cập nhật mã, khởi động lại GUI và xuất lại Excel trước khi xuất PDF;
không dùng đợt Excel cũ có bố cục 68%. Các ghi chép chưa chạy test dưới
đây mô tả lần bàn giao trước, không phải kết quả chẩn đoán mới.

Đã kiểm tra cú pháp tĩnh. Đã cập nhật/bổ sung test cho luồng đọc song song,
thứ tự kết quả, loại trừ mẫu/kết quả và xuất một hộ chính thức, nhưng
**chưa chạy test, benchmark hoặc xuất dữ liệu** theo yêu cầu người dùng.
Chưa có số liệu để cam kết mức tăng tốc thực tế hoặc kết luận nghiệm thu.
