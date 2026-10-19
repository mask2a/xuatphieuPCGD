# PCGD_HOGIADINH — Trạng thái dự án

Cập nhật: 2026-10-19

## Giai đoạn hiện tại

**Giai đoạn 3 — Xuất Excel theo hộ: exporter đã triển khai và kiểm thử bằng
Excel COM thật; ánh xạ nghiệp vụ chưa hoàn tất.**

Giai đoạn 1 và 2 đã hoàn tất. Giai đoạn 3A có kiểm thử tổng hợp cho hộ
1/6/7/14/15 người trên mẫu thật. Chưa cho phép coi toàn bộ nghiệp vụ Excel
đã nghiệm thu vì các ánh xạ chưa chốt và cấu hình địa bàn/năm học còn thiếu.
Không triển khai PDF hoặc GUI.

Đã tạo:

- `docs/DATA_ANALYSIS.md` — sheet/cột nguồn, thống kê dữ liệu mẫu và cảnh báo.
- `docs/TEMPLATE_MAPPING.md` — cấu trúc phiếu, vùng gộp, hàng thành viên,
  ánh xạ và thử nghiệm phiếu ngắn.
- `docs/MAPPING_AUDIT.md` — giải thích 15 cảnh báo kiểm thử, audit trường
  nguồn/đích ứng viên, phân loại trường và điểm cần xác nhận.
- `docs/PHASE3_FINAL_AUDIT.md` — kết quả audit Giai đoạn 3, độ phủ mapping,
  dữ liệu thật, cấu hình và cổng nghiệm thu còn chặn.
- `docs/SPEC.md` — mục tiêu, quy trình Excel → nghiệm thu toàn đợt → PDF →
  nghiệm thu → gộp, bố cục `output/`, chạy lại PDF, kiến trúc và kế hoạch
  kiểm thử.
- `reader.py` — đọc `.xls`/`.xlsx`, dò sheet/header, giữ giá trị thô và dạng
  hiển thị, ghi nguồn từng dòng, tiếp tục qua lỗi mỗi tệp.
- `normalizer.py` — chuẩn hóa tên/mã/ngày, giữ bản ghi thô và không điền giá
  trị thiếu.
- `household_grouper.py` — nhóm theo địa bàn + tệp + mã; cảnh báo mã lặp
  xuyên tệp và dữ liệu hộ không nhất quán; không bỏ thành viên trùng tên.
- `analysis.py`, `warning_report.py`, `settings.py` — điều phối phân tích,
  xuất báo cáo XLSX, tải cấu hình ngoài.
- `src/template_mapper.py` — tải và kiểm tra cấu hình ánh xạ đã xác minh.
- `src/excel_exporter.py` — một phiên Excel COM/đợt, tạo từng `.xlsx` từ
  bản sao mẫu, đọc/kiểm tra cấu hình xã và năm học trước khi tạo output,
  tiếp tục sau lỗi từng hộ, kiểm tra lại workbook và trả báo cáo thành
  công/thất bại/cảnh báo.
- `config/template_mapping.json` — 14 vị trí thành viên, các ô đích được
  phép ghi, cấu trúc gộp, vùng in và danh sách trường chưa ánh xạ (đã bổ
  sung diện ưu tiên J vào danh sách cảnh báo).
- `tests/test_excel_exporter.py` — kiểm thử tích hợp bằng Microsoft Excel
  thật, dùng dữ liệu tổng hợp và mẫu gốc chỉ đọc.
- `requirements.txt` — khai báo `pywin32`; môi trường hiện đã cài pywin32
  312 cho Python 3.15.
- `config/settings.json` — cột nguồn, tên sheet, xã/năm học trống để
  người dùng cấu hình; không hard-code xã hay năm học.
- `tests/` — test reader, normalizer, grouper, báo cáo và Excel exporter.

Đã cập nhật:

- `PROJECT_STATUS.md` — trạng thái, kết quả đã kiểm tra và việc còn mở.

## Bằng chứng đã kiểm tra

- Đọc `dulieumau.xls` và `maubaocao.xls` bằng Excel COM 16.0, không ghi vào
  tệp mẫu.
- Nguồn có 427 bản ghi thành viên, 120 mã phiếu không rỗng; có 1 bản ghi
  thiếu năm sinh, 3 mã có biến thể tên chủ hộ, 28 mã có nhiều số giữa các
  thành viên và 5 mã có nhiều giá trị diện cư trú.
- Mẫu có 14 khối thành viên, 442 vùng gộp, phần xác nhận ở hàng 82–87.
- Thử trên bản sao mẫu: hộ 1 người và 6 người đều tạo PDF một trang; vùng
  gộp và nội dung xác nhận còn nguyên. Tệp mẫu gốc không bị sửa.
- Đây chỉ là thử nghiệm bố cục mẫu; **không** phải nghiệm thu Excel→PDF cho
  ứng dụng và không chứng minh các trường hợp 7/14 người hoặc gộp PDF.
- Giai đoạn 3: 5 workbook `.xlsx` tổng hợp (1, 6, 7, 14, 15 người) được
  xuất từ bản sao `maubaocao.xls`, mở lại bằng Excel, kiểm tra thành viên
  liên tục, 442 vùng gộp, định dạng, kích thước hàng/cột, thiết lập in,
  vùng xác nhận và external links. Kiểm tra cả năm sinh thiếu, số chủ hộ
  thiếu, nhiều số thành viên, tên chủ hộ/diện cư trú mâu thuẫn, tên file
  trùng sau chuẩn hóa và không ghi đè.
- Phân tích lại nguồn thật ở chế độ chỉ đọc: 1 tệp, 427 bản ghi, 120 hộ;
  3 hộ xung đột tên chủ hộ, 5 hộ xung đột diện cư trú, 10 hộ không xác định
  chủ hộ duy nhất. Không xuất hộ thật vì cấu hình xã/năm học còn thiếu.
- Hash của `maubaocao.xls` và `dulieumau.xls` không đổi trong test.
- Bản thử có thể mở lại tại `output/excel/kiem_thu_giai_doan_3/`; báo cáo
  đối chiếu đặt cạnh các workbook. Đây là dữ liệu tổng hợp, không phải phiếu
  hộ thực tế.

## Lộ trình cập nhật theo yêu cầu Excel/PDF

| Chặng | Phạm vi | Trạng thái |
|---|---|---|
| 1. Phân tích dữ liệu/mẫu | Đọc workbook, cấu trúc và thử bố cục phiếu ngắn | Hoàn tất |
| 2. Reader và nhóm hộ | Đọc `.xls`/`.xlsx`, giữ dữ liệu thô, chuẩn hóa khóa, cảnh báo trùng/xung đột và báo cáo | Hoàn tất; 14/14 test PASS |
| 3A. Xuất Excel | Một `.xlsx`/hộ; phiếu ngắn/đầy đủ/bổ sung | Đã triển khai; test tổng hợp 1/6/7/14/15 PASS |
| 3B. Nghiệm thu Excel | Mở lại file, xác minh thành viên, đánh số, cấu trúc, merge, vùng in, xác nhận và định dạng | 21/21 test PASS; chỉ fixture tổng hợp, nghiệm thu nghiệp vụ còn chặn do cấu hình và mapping chưa chốt |
| 4. Chuyển PDF | Excel COM; một phiên/đợt; có thể chạy lại từ Excel sẵn | Chưa làm; chỉ nhận Excel đã PASS |
| 5. Gộp và GUI | Gộp PDF hợp lệ, progress/log/thống kê, không ghi đè tổng hợp | Chưa làm |
| 6. Nghiệm thu/đóng gói | Excel + PDF Windows thật, các trường hợp 1/6/7/14/>14 và đóng gói | Chưa làm |

Không được bắt đầu chặng 4 cho tới khi lượt xuất Excel hoàn tất và toàn bộ
workbook đủ điều kiện đã qua chặng 3B. Có thể chạy lại chặng 4 từ Excel đã
có, nhưng phải nghiệm thu các file được chọn trước khi chuyển PDF.

## Đầu ra và thống kê dự kiến

```text
output/
├── excel/kiem_thu_giai_doan_3/  # 5 workbook test tổng hợp + báo cáo đối chiếu
├── pdf/                         # Chưa triển khai
├── TongHop_<ten_nguon>.pdf      # Chưa triển khai
└── xu_ly.log                    # Chưa triển khai
```

Giao diện sẽ mặc định bật ba lựa chọn xuất Excel, xuất PDF và gộp PDF. Phải
hiển thị riêng số Excel/PDF thành công, thất bại và số trang tổng hợp; ghi
nhật ký danh sách hộ đã gộp, số trang từng hộ và số hộ lỗi. Chưa có PDF nào
được tạo.

## Giai đoạn 2 — Quy tắc điện thoại và kết quả

Tên xã dùng khóa cấu hình ngoài `commune_name`; địa bàn và năm học không
được hard-code.

Điện thoại hộ chỉ lấy từ `AW` của thành viên duy nhất được nhận diện là chủ
hộ qua `AU`. Không tìm số thay thế ở thành viên khác. Chủ hộ không có số thì
để trống; thiếu hoặc nhiều chủ hộ thì để trống và cảnh báo; nhiều số trên
các bản ghi chủ hộ cũng được cảnh báo. Nhiều số của thành viên khác nhau
không bị coi là lỗi hộ.

Đã xác minh lại trên `dulieumau.xls`: 427 bản ghi, 120 nhóm; 1 năm sinh
trống; 3 mã cảnh báo biến thể tên chủ hộ; 5 mã cảnh báo biến thể diện cư trú;
28 mã có nhiều số điện thoại giữa các thành viên nhưng không tạo cảnh báo
xung đột điện thoại hộ. Các bản ghi chủ hộ không đủ căn cứ không được tự
chọn số.

Đã tạo file `output/CANH_BAO_DU_LIEU.xlsx` từ workbook mẫu, gồm sheet tổng
quan, danh sách 120 hộ/số thành viên và các cảnh báo theo hộ/nguồn. Báo cáo
ghi nhận 18 cảnh báo: 10 hộ không xác định được chủ hộ từ quan hệ, 3 mã có
biến thể tên chủ hộ, 5 mã có biến thể diện cư trú; không có cảnh báo xung
đột điện thoại chủ hộ.

Lệnh test giai đoạn 2: `.\venv\Scripts\python.exe -m unittest discover -v`.
Kết quả lịch sử: **14 test PASS, 0 FAIL**. Phạm vi đó kiểm thử reader,
normalizer, grouper và báo cáo, không phải nghiệm thu Excel/PDF.

## Giai đoạn 3 — Kết quả kiểm thử Excel

Đã chạy lại toàn bộ suite bằng lệnh:
`.\venv\Scripts\python.exe -m unittest discover -v` — **21 PASS, 0 FAIL**.
Mười chín test trước tiếp tục PASS; thêm hai test preflight cấu hình. Các
test Excel dùng Microsoft Excel COM thật trên Windows và `maubaocao.xls`
chỉ đọc.

Đã tạo 5 workbook kiểm thử tổng hợp cho 1, 6, 7, 14 và 15 thành viên cùng
báo cáo đối chiếu. Báo cáo xác nhận số người/thứ tự, 442 vùng gộp mỗi sheet,
phần xác nhận, vùng in, external links và bất biến hash của mẫu/nguồn.
Chi tiết nằm trong `output/excel/kiem_thu_giai_doan_3/BAO_CAO_DOI_CHIEU.json`.

Đã audit 15 lần cảnh báo ở 5 workbook: mỗi workbook có cảnh báo diện cư trú,
năm học chưa ánh xạ và giới tính chưa ánh xạ. Đã thêm diện ưu tiên J vào
danh sách cảnh báo trường chưa ánh xạ. Không tìm thấy trường mới đủ căn cứ
để ghi; tổng mapping hiện có 10 quy tắc đích/13 cột nguồn, với 36 trường
nghiệp vụ còn khai báo chưa ánh xạ.

## Việc chưa xử lý / cần xác nhận

1. `config/settings.json` vẫn có `commune_name=""`, `school_year=null`.
   Exporter hiện chặn cả xã/năm học thiếu, năm học sai định dạng hoặc địa bàn
   hộ không khớp cấu hình trước khi tạo thư mục đầu ra/khởi chạy COM.
2. Mẫu còn in sẵn các năm 2025–2030. Vì quy tắc dòng theo năm học chưa được
   xác nhận, exporter giữ nguyên nội dung mẫu và cảnh báo; không tự suy ra
   năm học từ `school_year`.
3. Chưa ghi các trường J, M/O/G, P, Q:T, V:AH, AI:AR, AS:AX (ngoại trừ AU,
   AV, AW) khi chưa có quy tắc xác nhận/ô đích. Báo cáo exporter liệt kê
   cảnh báo cho các trường nguồn có dữ liệu.
4. Ký tự `q` trong checkbox là font Wingdings; ký tự thay thế để đánh dấu
   giới tính/diện cư trú chưa được xác minh. Các checkbox được giữ nguyên.
5. Kết quả xuất thử hiện dùng dữ liệu tổng hợp: 5 thành công, 0 thất bại.
   Chưa xuất `.xlsx` hộ thật cho 120 hộ vì cấu hình chưa có giá trị hợp lệ;
   dữ liệu nguồn không thay đổi.
6. PDF, gộp PDF, GUI và EXE chưa triển khai; không có test nào của các chặng
   đó được coi là PASS.

## Bước tiếp theo

Tiếp theo cần tên xã/năm học thực tế và quy tắc còn mở ở
`docs/MAPPING_AUDIT.md`. Sau đó bổ sung test cho từng mapping, xuất thử các
hộ nguồn đại diện (bao gồm hộ mâu thuẫn), đối chiếu Excel từng trường và
hoàn tất cổng nghiệm thu Excel. Chỉ sau khi cổng Excel nghiệp vụ PASS mới
bắt đầu Excel COM → PDF, xác thực PDF, gộp PDF và kiểm thử chạy lại từ Excel
sẵn có.
