# Phân trang 6 người, in 68% — 2026-10-21

> Cập nhật sau báo cáo: người dùng tự sửa độ rộng cột trong mẫu và yêu
> cầu giảm font phần Dân tộc xuống 8 pt trên bản xuất. Mã đã cập nhật;
> test đang chạy đã dừng theo yêu cầu. Các PDF/số đo dưới đây thuộc bản
> trước thay đổi font và lần sửa mẫu mới nhất, không phải bản nghiệm thu
> của mã mới. Xem PROJECT_STATUS.md để biết trạng thái hiện hành.

Quy tắc hiện hành theo xác nhận người dùng: **6 người/trang, giữ 68% nếu
vừa**. Thay thế quy tắc 7 hoặc 14 người/trang và việc chuẩn hóa kích thước
cột/dòng ở lượt trước.

## Cách triển khai

- Mỗi nhóm tối đa 6 người dùng một bản sao đầy đủ của sheet mẫu; dùng sáu
  vị trí đã có trong mapping: dòng 9, 14, 19, 24, 29, 34. Ánh xạ cột và
  quy tắc nghiệp vụ không đổi. Số thứ tự tiếp tục xuyên suốt các sheet.
- Giữ nguyên sáu khối, cả các khối chưa có người. Chỉ ẩn hàng 39:81 để
  phần đầu phiếu 1:8, sáu khối 9:38 và vùng ký 82:87 cùng xuất hiện.
  Không xóa hàng, khối hoặc vùng gộp; không thay ColumnWidth/RowHeight.
- Bỏ nhân độ rộng cột 1,7 và chỉnh chiều cao dòng của lượt chuẩn hóa cũ.
  Không tự đổi font, WrapText hoặc ShrinkToFit khi gặp nội dung dài.
  Ngoại lệ font duy nhất là các vị trí checkbox Unicode đã được duyệt:
  ☒/☐ với Segoe UI Symbol. Giữ nguyên các định dạng ô khác.
- PageSetup dùng chung Excel/PDF: A4 ngang, PrintArea=$A$1:$AG$87,
  **Zoom=68**; bốn lề 0,5 cm; header/footer 0; căn giữa ngang/dọc.
  FitToPagesWide/Tall vẫn lưu 1 nhưng không chi phối khi Zoom=68.
- Mỗi sheet phải xuất đúng một trang PDF A4 ngang. Đối chiếu nội dung
  theo đúng trang và thứ tự thành viên; mỗi trang phải có đầu phiếu và
  vùng ký. Các trường bị mất/cắt trong văn bản PDF sẽ chặn nhận PDF.
- Ô có nội dung dài được cảnh báo LONG_TEXT_REQUIRES_REVIEW, không tự co
  chữ để lách quy tắc giữ định dạng mẫu. Vẫn cần rà soát độ dễ đọc của các
  chuỗi thực tế quá dài; không tuyên bố mọi chuỗi dài tùy ý đều in vừa.

## Kiểm tra bảo toàn định dạng

Exporter mở lại file trước khi nhận kết quả: kiểm tra tất cả độ rộng cột,
chiều cao dòng lưu trong XML (cả hàng ẩn), trạng thái ẩn phân trang, 442
vùng gộp, font/đường viền/căn lề tại các ô kiểm chứng, giá trị đã ánh xạ,
checkbox và PageSetup. COM trả chiều cao 0 cho hàng ẩn nên không dùng
giá trị đó để kết luận mất chiều cao gốc.

Đã kiểm chứng PDF mẫu chỉ đọc: 7 khối ở 68% làm vùng ký tràn sang trang 2;
6 khối ở 68%, lề 0,5 cm vừa một trang. Không lưu thay đổi vào mẫu gốc.

| Số người | Số sheet/trang PDF yêu cầu |
| --- | --- |
| 1 | 1 |
| 7 | 2 |
| 8 | 2 |
| 14 | 3 |
| 15 | 3 |
| 21 | 4 |

## Kết quả bộ xuất bàn giao

Thư mục `output/testing/gui_batch_pdf_20261021_024156/`: 6 Excel, 6 PDF
riêng, 0 lỗi, PDF tổng hợp 15 trang đúng bảng trên. Tất cả A4 ngang, mỗi
sheet đúng một trang. Đã xem đủ 15 ảnh render trong `previews/`: đủ đầu
phiếu, khối trống, năm học, vùng ký, checkbox và các trường giả lập.

Đo bằng PyMuPDF: bề ngang bảng chiếm 97,55% vùng in sau lề, lệch tâm
0,36 pt; chữ tên nguồn 10,5 pt giữ nguyên trong Excel, ra PDF khoảng
7,2 pt đồng đều giữa các trang. `VERIFICATION.json` lưu số đo từng trang
và xác nhận hash nguồn/mẫu/mapping trước/sau không đổi.
Đã render cả 15 trang PDF tổng hợp và so pixel với từng trang PDF riêng
đã xem: tất cả khớp hoàn toàn, không thay đổi hình thức sau gộp.

PDF tổng hợp:
`output/testing/gui_batch_pdf_20261021_024156/pdf/dot_23d50b21908a4ded934aa8cdd162003a/TongHop.pdf`.

Thứ tự text nội bộ PDF Excel có thể nhóm theo font/màu; kiểm tra thứ tự
thành viên phải dùng chế độ trích xuất theo bố cục trang của pypdf, không
dùng thứ tự thô của các lệnh vẽ. Có test chặn mất tên, đảo tên, mất trường
và mất vùng ký; đối chiếu cả số thứ tự/tên trong từng sheet Excel.

Kết quả test đầy đủ được ghi trong PROJECT_STATUS.md sau khi hoàn tất.
Script tái tạo: `scripts/verify_gui_batch_pdf.py`;
dữ liệu hoàn toàn giả lập, gồm lớp/trường/tốt nghiệp/tôn giáo và các
trường bổ sung. Các bộ PDF cũ được giữ để đối chiếu lịch sử.

Có thể chạy lại bước PDF từ bộ Excel giả lập đã tạo bằng
`scripts/verify_gui_batch_pdf.py --reuse-excel output/testing/gui_batch_pdf_20261021_024156`.
Tất cả Excel vẫn được nghiệm thu lại bằng COM trước khi chuyển PDF;
mỗi lần PDF dùng thư mục mới, không ghi đè PDF cũ.

File sửa: `src/print_layout.py`, `src/excel_exporter.py`,
`src/export_workflow.py`, các test tương ứng, script kiểm chứng và tài liệu.
Không sửa `dulieumau.xls`, `maubaocao.xls`, mapping hoặc snapshot. Không commit.
