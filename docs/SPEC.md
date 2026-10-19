# Đặc tả phần mềm PCGD_HOGIADINH

## Mục tiêu

Đọc một hoặc nhiều workbook Excel phổ cập giáo dục, bảo toàn dữ liệu nguồn,
nhóm thành viên theo hộ, và xuất một workbook `.xlsx` cho mỗi hộ dựa trên
mẫu Excel duy nhất `maubaocao.xls`.

## Phạm vi chức năng

- Đọc `.xls` và `.xlsx` từ một tệp hoặc một thư mục.
- Bỏ qua tệp tạm, tệp mẫu và tệp kết quả; lỗi một tệp không được làm dừng
  cả đợt.
- Thống kê hộ/nhân khẩu, báo lỗi thiếu mã và cảnh báo xung đột.
- Phát hiện mã phiếu trùng giữa tệp; không gộp tự động chỉ dựa trên mã.
- Xuất một tệp cho mỗi hộ, không ghi đè lên tệp có sẵn, đóng workbook sau
  mỗi thao tác và tiếp tục hộ khác nếu một hộ bị lỗi.
- Giai đoạn xuất Excel dùng duy nhất `maubaocao.xls`, `win32com.client`,
  một phiên Excel COM cho mỗi đợt, và lưu một `.xlsx` cho mỗi hộ vào
  `output/excel/`. Tên `ho_<ma_ho>.xlsx` được chuẩn hóa theo quy tắc tên tệp
  Windows; kết quả cũ không bị ghi đè.
- Excel exporter trả báo cáo có số file thành công/thất bại và cảnh báo theo
  hộ. Mỗi file được mở lại bằng Excel để kiểm tra trước khi tính là thành
  công; lỗi một hộ không chặn hộ tiếp theo.
- Xuất Excel cho từng hộ trước; chỉ chuyển sang PDF sau khi tệp Excel đó đã
  được tạo và vượt qua kiểm tra nghiệm thu tự động. Không chuyển hộ Excel
  thất bại hoặc không hợp lệ sang bước PDF.
- Dùng Excel COM/pywin32 để sao chép mẫu `.xls`, lưu `.xlsx`, rồi xuất mỗi
  Excel hợp lệ thành PDF bằng `win32com.client` và `ExportAsFixedFormat`.
  Không sửa mẫu gốc; duy trì định dạng, ô gộp, thiết lập in, tỷ lệ và ngắt
  trang; chỉ in vùng hợp lệ; không để mất phần xác nhận hoặc tạo trang trắng.
- Một hộ xuất PDF lỗi không làm dừng các hộ khác. Tái sử dụng một phiên
  Excel COM trong mỗi đợt; luôn đóng workbook và giải phóng COM resources,
  kể cả khi một hộ gặp lỗi.
- Cho phép chạy lại riêng bước Excel → PDF từ các `.xlsx` đã xuất mà không
  cần đọc lại dữ liệu nguồn. Tệp Excel đưa vào vẫn phải qua kiểm tra hợp lệ.
- Gộp các PDF đã xuất thành công và đã được xác thực bằng `pypdf` hoặc
  `PyPDF2`, theo thứ tự mã hộ xác định. Không đưa PDF lỗi, dở dang hoặc thiếu
  trang vào bản tổng hợp; ghi nhận mã hộ, số trang và số hộ lỗi.
- Không ghi đè PDF tổng hợp đã tồn tại nếu chưa được người dùng cho phép.
- Giao diện Tkinter/ttk có chọn tệp/thư mục, phân tích, xuất, tiến trình,
  thống kê, nhật ký, mở thư mục kết quả; tác vụ dài chạy nền. Có ba lựa chọn
  riêng: xuất Excel từng hộ, xuất PDF từng hộ, gộp PDF tổng hợp; mặc định
  bật cả ba. Có thể chạy lại PDF từ thư mục Excel sẵn có.
- Cấu hình JSON cho ánh xạ cột nguồn, ô đích, sheet và các hàng thành viên.

## Quy trình xử lý và cổng nghiệm thu

Quy trình bắt buộc:

`Dữ liệu nguồn → Nhóm hộ → Xuất Excel từng hộ → Nghiệm thu Excel → Chuyển PDF từng hộ → Nghiệm thu PDF → Gộp PDF tổng hợp`

Mỗi hộ được xử lý độc lập. Khi xuất Excel được bật, hoàn tất lượt xuất và
nghiệm thu tất cả workbook trước khi bắt đầu bất kỳ lượt chuyển PDF nào.
Workbook của hộ bị lỗi không chặn các hộ khác nhưng không được chuyển PDF.
Nghiệm thu Excel phải kiểm tra workbook mở lại được, đúng số thành viên và
đánh số liên tục, đúng loại phiếu (ngắn/đầy đủ/bổ sung), giữ các ô gộp và
phần xác nhận, cùng vùng in/thiết lập trang hợp lệ. Chỉ workbook qua cổng
này mới đủ điều kiện chuyển PDF.

Trong chế độ chạy lại PDF, lấy `.xlsx` có sẵn từ `output/excel/` (hoặc thư
mục Excel người dùng chọn), không đọc lại nguồn. Trước khi xuất, vẫn nghiệm
thu cấu trúc workbook. Bỏ qua và ghi lỗi các Excel không mở được hoặc không
đạt kiểm tra; tiếp tục các Excel khác.

Chỉ PDF xuất hoàn tất và mở/đọc được mới đủ điều kiện gộp. Gộp PDF theo mã hộ
tăng dần theo giá trị mã gốc, không sửa mã khi sắp xếp; mã trùng/xung đột giữa
các nguồn vẫn theo quy tắc phát hiện trùng, không âm thầm gộp hộ. Ghi số
trang của mỗi hộ, tổng số hộ/PDF lỗi và tổng số trang. Tạo bản tổng hợp tạm,
xác thực lại, rồi mới đưa ra tên cuối;
không thay thế bản cũ nếu chưa có xác nhận cho phép ghi đè. Một file PDF lỗi
không được làm dừng chuyển đổi các hộ khác.

## Thư mục đầu ra

```text
output/
├── excel/                    # Một .xlsx cho mỗi hộ
├── pdf/                      # Một .pdf tương ứng mỗi Excel hợp lệ
├── TongHop_<ten_nguon>.pdf   # PDF tổng hợp của đợt
└── xu_ly.log                 # Kết quả, lỗi và thống kê từng bước
```

Tên tệp nguồn trong tên PDF tổng hợp phải được làm an toàn cho tên tệp trên
Windows. Không ghi đè Excel/PDF hộ đã tồn tại; xung đột tên phải được báo cáo
thay vì âm thầm thay thế. Ghi log cả các hộ không gộp, số trang mỗi PDF được
gộp, tổng số trang và số hộ thất bại ở từng bước.

## Giao diện và tiến trình

- Ba checkbox bật mặc định: **Xuất Excel từng hộ**, **Xuất PDF từng hộ**,
  **Gộp PDF tổng hợp**.
- Nếu PDF được chọn nhưng không chọn xuất Excel trong đợt hiện tại, chạy từ
  các Excel sẵn có đã được chọn/đọc trong thư mục Excel kết quả; không yêu
  cầu đọc lại dữ liệu nguồn.
- Thanh tiến trình và nhật ký phải phân biệt giai đoạn Excel, PDF và gộp.
- Thống kê riêng số Excel thành công/thất bại, PDF thành công/thất bại, số
  trang bản tổng hợp; báo lỗi mà không chặn các hộ độc lập khác.
- Hủy/đóng cửa sổ không được làm Excel COM hoặc workbook bỏ mở.

## Nguyên tắc dữ liệu

- Đọc sheet chính `MauNhapLieu`, dữ liệu từ hàng 5; tiêu đề nằm ở hàng 2 và
  một số tiêu đề con ở hàng 3.
- Giữ nguyên giá trị gốc, bao gồm số 0 đầu, mã dạng chuỗi, ngày/tháng và
  khoảng trắng. Không thay trống bằng giá trị mặc định.
- Giữ giá trị thô và tách riêng khóa nhóm đã chuẩn hóa nếu cần; không ghi
  ngược giá trị đã chuẩn hóa vào dữ liệu nguồn.
- Mã `N` không được dùng đơn độc để gộp hộ trên nhiều tệp. Mã trùng/xung đột
  phải được báo cáo; cần danh tính nguồn và địa bàn trước khi quyết định gộp.
- Nếu tên chủ hộ khai báo hoặc diện cư trú có nhiều giá trị, ghi cảnh báo và
  không tự chọn một giá trị.
- Điện thoại hộ chỉ lấy từ cột `AW` của bản ghi duy nhất có quan hệ chủ hộ
  xác định qua `AU`. Không dùng điện thoại thành viên khác làm giá trị thay
  thế; nếu chủ hộ duy nhất thiếu số thì để trống. Nếu không xác định duy nhất
  chủ hộ hoặc nhiều bản ghi chủ hộ có số khác nhau, để trống và cảnh báo.
  Nhiều số điện thoại của các thành viên không phải chủ hộ không phải lỗi.
- Tên xã/địa bàn và năm học chỉ lấy từ `config/settings.json` hoặc cấu hình
  được cung cấp bên ngoài; không mã hóa sẵn `VĂN HÁN` hay năm học trong code.
- Trước khi xuất Excel, bắt buộc kiểm tra `commune_name` không rỗng và
  `school_year` có dạng `YYYY-YYYY` với hai năm liên tiếp; địa bàn của từng
  hộ phải khớp với cấu hình. Nếu thiếu/sai, dừng trước khi tạo thư mục đầu
  ra hoặc khởi chạy Excel COM. Năm học cấu hình không tự ghi vào các dòng
  năm có sẵn của mẫu khi quy tắc ánh xạ dòng chưa được xác nhận.
- Tạo `output/CANH_BAO_DU_LIEU.xlsx` với số tệp đọc/lỗi, số bản ghi, tổng số
  hộ, số thành viên từng hộ và danh sách các hộ có cảnh báo/mâu thuẫn.

## Quy tắc mẫu đã xác minh

- Sheet: `Phiếu điều tra 1 mặt`; mẫu có 14 vị trí, 5 hàng mỗi vị trí.
- Hộ 1–6: ẩn các hàng thành viên chưa dùng, không xóa hàng; vùng in
  `$A$1:$AG$87`, giữ phần xác nhận hàng 82–87.
- Hộ 7–14: dùng cấu trúc đầy đủ.
- Trên 14: thêm sheet bổ sung trong cùng workbook và đánh số liên tục; thiết
  kế và kiểm thử chi tiết trong giai đoạn xuất.
- Phiếu ngắn 1–6 người: ẩn phần hàng thành viên chưa dùng, không xóa hàng;
  đặt PrintArea `$A$1:$AG$87` và Fit-to-Page một trang.
- Bảng ánh xạ được phép dùng và các trường chưa rõ nằm trong
  `TEMPLATE_MAPPING.md` và `config/template_mapping.json`. Không xem ánh xạ
  "ứng viên" là quy tắc đã được phê duyệt. Trường chưa chốt giữ nguyên mẫu
  và tạo cảnh báo, không tự đoán.

## Kiến trúc mục tiêu

```text
main.py
├── gui.py                    # Tkinter; gửi lệnh nền, nhận tiến trình/kết quả
├── workflow.py               # Điều phối một đợt và thứ tự Excel → PDF → gộp
├── excel_com_session.py      # Một phiên Excel sở hữu bởi thread xử lý mỗi đợt
├── reader.py                 # Đọc workbook nguồn; không làm biến đổi dữ liệu
├── normalizer.py             # Giá trị thô + khóa chuẩn hóa, không điền ô trống
├── household_grouper.py      # Nhóm/cảnh báo xung đột và mã trùng
├── analysis.py               # Kết quả phân tích dùng chung cho báo cáo
├── warning_report.py         # Báo cáo Excel số hộ, thành viên và cảnh báo
├── settings.py               # Cấu hình địa bàn, năm học, cột và đường dẫn
├── template_analyzer.py      # Kiểm tra sheet, khối, merge và vùng in mẫu
├── src/template_mapper.py    # Ánh xạ đã được xác nhận từ JSON
├── src/excel_exporter.py     # Một phiên COM; tạo .xlsx từ bản sao mẫu
├── excel_validator.py        # Cổng nghiệm thu workbook trước PDF
├── pdf_exporter.py           # Adapter Excel COM: .xlsx hợp lệ → PDF
├── pdf_validator.py          # Kiểm tra PDF đọc được, trang và tính toàn vẹn
├── pdf_merger.py             # Gộp PDF hợp lệ theo mã hộ; không ghi đè mặc định
└── logging/progress           # Nhật ký và thông báo tiến độ cho GUI
```

Giai đoạn 2 cung cấp các interface `read_sources(paths, settings)`,
`normalize_records(records, settings)`, `group_households(members)`,
`analyze_sources(paths, settings)` và `write_warning_report(result, path)`.
Mỗi bản ghi chuẩn hóa giữ nguyên `RawRecord` cùng tệp, sheet, số dòng nguồn.
Nhóm hộ cô lập theo mã, địa bàn cấu hình và tệp nguồn; mã lặp giữa các tệp
được cảnh báo, không gộp. Danh sách cột và đường dẫn cấu hình nằm trong
`config/settings.json`.

Giai đoạn 3 triển khai interface `export_households(households, template_path,
output_dir, mapping_path)` trong `src/excel_exporter.py`. Interface dùng một
phiên Excel COM, trả danh sách thành công/thất bại/cảnh báo và mở lại từng
workbook đã lưu để nghiệm thu trước khi báo thành công. Các ô nguồn/đích và
trường chưa ánh xạ được khai báo trong `config/template_mapping.json`.
Workbook không hợp lệ hoặc có đích trùng tên không được ghi đè.

`workflow.py` là module điều phối sâu với giao diện nhỏ: nhận yêu cầu chạy và
phát sự kiện tiến trình/kết quả. Nó giữ thứ tự, áp dụng cổng nghiệm thu toàn
đợt giữa Excel và PDF, tiếp tục hộ độc lập khi lỗi, và chỉ gộp tài liệu đã
được xác thực. `excel_com_session.py` mở đúng một phiên Excel cho một đợt và
được `workflow.py` dùng chung cho các thao tác xuất Excel và xuất PDF; toàn
bộ COM chạy tuần tự trên thread sở hữu phiên, không gọi từ GUI. Trong chế độ
chạy lại PDF, cùng phiên được dùng để nghiệm thu/mở và chuyển đổi các Excel
đã có; chế độ này bỏ qua reader/grouper/exporter nhưng không bỏ qua validator.

Các validator có thể kiểm tra workbook/PDF đã lưu qua interface tệp thực;
kiểm thử tích hợp Windows dùng Excel COM thật và các tệp đầu ra tạm. Không
để logic tạo file, COM lifecycle hay phục hồi lỗi rải rác trong GUI.

## Kế hoạch kiểm thử và nghiệm thu

**Cổng Excel (bắt buộc phải PASS trước PDF):**

1. Xuất hộ đúng 1, 6, 7 và 14 thành viên; xác minh hộ 1–6 dùng phiếu ngắn,
   hộ 7–14 dùng phiếu đầy đủ.
2. Mở lại từng `.xlsx` bằng Excel; so sánh danh sách/tổng số thành viên và
   số thứ tự với đầu vào; kiểm tra không mất/trùng người.
3. So sánh vùng gộp, vùng in, lề, khổ giấy, hướng giấy, tỷ lệ/fit và ngắt
   trang với mẫu. Kiểm tra xác nhận hàng 82–87 còn nguyên; không sửa mẫu gốc.
4. Kiểm tra trường hợp trên 14 người: tất cả thành viên xuất hiện liên tục
   trên sheet/trang bổ sung cùng workbook; không tạo thêm workbook hộ.
5. Lỗi lưu, workbook hỏng, sai số người, sai cấu trúc, hoặc sai thiết lập in
   phải đánh FAIL; không gọi PDF exporter cho workbook đó.

**Cổng PDF và gộp:**

1. Dùng Excel COM thật (`win32com.client`) để xuất các Excel đã PASS bằng
   `ExportAsFixedFormat`; xác minh chỉ vùng in dự kiến được xuất.
2. Kiểm tra số trang thực tế cho hộ 1, 6, 7, 14 và hộ trên 14; đối chiếu
   từng trang với bố cục kỳ vọng đã được xác nhận, không có trang trắng,
   không mất thành viên, trang bổ sung hay phần xác nhận cuối phiếu.
3. Mở/đọc PDF bằng `pypdf` hoặc `PyPDF2`; PDF lỗi, không đọc được hoặc chưa
   hoàn tất phải bị loại khỏi danh sách gộp.
4. Chạy lại riêng chuyển PDF từ Excel sẵn có; xác minh không cần nguồn và
   cùng số trang/kết quả như lần chuyển đầu.
5. Gộp nhiều PDF thành công; kiểm tra thứ tự mã hộ, đủ số hộ, tổng số trang
   bằng tổng trang thành phần, không thêm/trùng trang. Ghi số hộ lỗi và lý
   do vào log.
6. Kiểm tra lỗi một workbook/PDF không dừng hộ kế tiếp; kiểm tra workbook và
   phiên COM luôn được đóng/giải phóng.
7. Khi file tổng hợp đích đã tồn tại, phải giữ nguyên nếu chưa được cho phép
   ghi đè. Kiểm tra PDF tổng hợp tạm/hỏng không được xuất hiện dưới tên cuối.

Không đánh dấu các trường hợp trên PASS cho tới khi chạy kiểm thử trên
Windows có Microsoft Excel và lưu kết quả/fixture kiểm thử.

**Kiểm thử giai đoạn 2:** chạy `python -m unittest discover -v`; đối chiếu
tệp mẫu với 427 bản ghi và 120 mã hộ; xác minh 1 năm sinh trống, 3 mã có
biến thể tên chủ hộ, 28 mã có số điện thoại khác nhau giữa các thành viên
(không gắn nhầm thành lỗi hộ), 5 mã có biến thể diện cư trú; kiểm tra mã
trùng giữa tệp, vùng địa bàn khác nhau, thành viên trùng tên, lỗi tệp độc
lập, số 0 đầu và không biến đổi dữ liệu gốc. Các kiểm thử giai đoạn 3–6 ở
trên vẫn là các cổng riêng, không được coi là đã chạy trong giai đoạn 2.

## Ngoài phạm vi giai đoạn 1

Chưa viết module xử lý, giao diện, xuất hàng loạt, đóng gói EXE, hay chốt
quy tắc nghiệp vụ đang mở. Không dùng workbook phụ `12..10 - PHỔ CẬP TDP 18
(1).xls` để tính kết quả mẫu.
