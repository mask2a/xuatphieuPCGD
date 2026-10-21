# Audit nhãn năm học và ánh xạ dữ liệu theo năm

Ngày kiểm tra: 2026-10-20
Mẫu kiểm tra: `maubaocao.xls`, mở bằng Microsoft Excel COM ở chế độ chỉ đọc.

## Kết luận

Người dùng đã phê duyệt **Phương án A – Năm học linh hoạt**: các nhãn năm
học trên bản sao biểu mẫu sẽ phụ thuộc `school_year`. Kiểm tra trực tiếp
cho thấy vùng nhãn lặp lại và vị trí các ô có thể xác định được. Tuy nhiên,
mẫu chỉ cung cấp một dãy cố định (2025–2026 đến 2029–2030). Người dùng đã
duyệt điểm neo tại **dòng nhãn thứ nhất** và năm lịch ở phần ký phải giữ
nguyên. Quy tắc được triển khai trong module riêng để cập nhật bản sao;
không nối vào luồng xuất hộ chính thức và không ánh xạ dữ liệu B5–B12.

## 1. Bằng chứng Excel COM về các nhãn năm

Workbook có một sheet `Phiếu điều tra 1 mặt`, `UsedRange` là `$A$1:$AG$94`.
Đã quét các ô có giá trị trong UsedRange; các nhãn năm học nằm ở cột F,
mỗi ô gộp ngang F:G.

| Vị trí thành viên | Các ô nhãn năm | Giá trị đang lưu theo thứ tự hàng |
|---:|---|---|
| 1 | `F9:F13` (mỗi ô gộp với cột G) | 2025 - 2026; 2026 - 2027; 2027 - 2028; 2028 - 2029; 2029 - 2030 |
| 2 | `F14:F18` | như trên |
| 3 | `F19:F23` | như trên |
| 4 | `F24:F28` | như trên |
| 5 | `F29:F33` | như trên |
| 6 | `F34:F38` | như trên |
| 7 | `F39:F43` | như trên |
| 8 | `F47:F51` | như trên |
| 9 | `F52:F56` | như trên |
| 10 | `F57:F61` | như trên |
| 11 | `F62:F66` | như trên |
| 12 | `F67:F71` | như trên |
| 13 | `F72:F76` | như trên |
| 14 | `F77:F81` | như trên |

Tổng cộng **70 ô nhãn năm học**: 14 vị trí thành viên × 5 hàng. Khoảng
hàng 44–46 không có khối thành viên. Từng ô năm là chuỗi văn bản ở ô neo
cột F; vùng gộp là `F{row}:G{row}`. Mỗi khối lặp lại cùng chuỗi năm theo
hàng, không thấy một bố cục khác ở các vị trí 8–14.

### Định dạng và cấu trúc liên quan

- Các ô nhãn năm có font Arial 10, đậm, không nghiêng, định dạng số `@`,
  bật WrapText; căn ngang trái và theo chiều dọc như nhau. Màu nền trắng.
- Các đường viền không hoàn toàn đồng nhất: kiểu biên trên/dưới thay đổi
  theo đầu/cuối các đoạn và vị trí khối. Không được tái tạo/ghi đè định
  dạng; nếu cập nhật giá trị trong tương lai chỉ sửa nội dung ô neo F, giữ
  nguyên merge và format.
- Chiều cao các dòng có khác biệt theo vị trí: những hàng đầu ở các vị trí
  thành viên muộn có thể cao 18.75, trong khi nhiều hàng là 18.0. Không
  giả định chiều cao các khối giống nhau.
- Chiều rộng cột A:AG đọc được lần lượt là: `2.86, 6.14, 4.14, 9.29,
  11.00, 4.29, 6.29, 4.43, 4.29, 4.43, 4.29, 4.71, 5.86, 4.29, 3.14,
  7.00, 6.14, 7.00, 6.29, 3.57, 3.57, 5.57, 5.43, 5.71, 6.14, 4.71,
  6.00, 4.71, 5.57, 6.00, 6.00, 12.71, 6.00`.
- Mẫu có 442 vùng gộp. Trang được đặt A4 ngang, `Zoom=68`,
  `FitToPagesWide=1`, `FitToPagesTall=1`; PrintArea trong workbook gốc đang
  rỗng. Excel COM báo hai ngắt trang ngang tại `$A$45`, `$A$92`, một ngắt
  dọc tại `$AH$1`.
- Hàng xác nhận bắt đầu ở 82. Các vùng ngày ký gộp là
  `F82:I82`, `J82:M82`, `N82:Q82`, `R82:V82`, `W82:Z82`; chúng ghi lần lượt
  `Năm 2025`, `Năm 2026`, `Năm 2027`, `Năm 2028`, `Năm 2029`.
  `AA82:AG83` chứa nội dung xác nhận của UBND và placeholder
  `Ngày … tháng … năm 20…`.

## 2. Năm không phải nhãn năm học

Các năm ở vùng xác nhận `F82:Z82` là năm lịch gắn với chỗ ký/ngày tháng,
không phải dãy năm học ở cột F của khối thành viên. Placeholder `20……`
trong nội dung xác nhận cũng là một thành phần năm lịch chưa điền.

Không có ô số/năm Excel dạng số nào được tìm thấy trong UsedRange. Ngoài 70
nhãn năm học và các năm lịch ở phần xác nhận, không tìm thấy giá trị năm
độc lập khác trong các ô có nội dung. Không được tự thay các năm lịch chữ
ký theo `school_year`.

## 3. Mối liên hệ giữa nhãn năm và các trường giáo dục

Hàng tiêu đề trên mẫu xác định vùng theo năm như sau:

- Hàng 4: `F` — “NHÓM, LỚP ĐANG HỌC (Theo năm học)”; `K` — “TÊN TRƯỜNG -
  HUYỆN ĐANG HỌC (Tương ứng với từng năm học)”; nhóm thông tin tốt nghiệp
  bắt đầu ở P.
- Hàng 6: P “Cấp học”, Q “Bổ túc”, R “Năm học”; S “Bậc”, T “Năm học”;
  V “Lớp”, W “Năm học”; X “Lớp”, Y “Năm học”; Z “Đang học lớp”, AA “Hoàn
  thành lớp”, AB “Tái mù chữ mức”; AC “Loại”, AD “Có CNKT”, AE “KNHT”.
- Hàng 8 đánh số cột/bảng, không cung cấp quy tắc quy đổi năm nguồn thành
  dòng năm trên mẫu.

Nhãn năm học lặp theo từng hàng thành viên; các cột dữ liệu giáo dục ở
cùng hàng là vị trí ứng viên tương ứng. Bằng chứng cấu trúc này **không**
xác định trường nguồn nào được ghi vào dòng nào, hoặc năm sự kiện trong
nguồn có phải năm học hay năm dương lịch.

## 4. Cơ chế linh hoạt đã được duyệt

Cấu hình `school_year` được kiểm tra theo mẫu `YYYY-YYYY` với hai năm liên
tiếp. Người dùng đã phê duyệt năm cấu hình là nhãn **đầu tiên** của từng
khối. Với năm bắt đầu Y, dòng i từ 0 đến 4 hiển thị `Y+i` đến `Y+i+1`.

| Cấu hình | 5 nhãn theo thứ tự dòng trong mỗi khối |
|---|---|
| `2026-2027` | `2026-2027`, `2027-2028`, `2028-2029`, `2029-2030`, `2030-2031` |
| `2027-2028` | `2027-2028`, `2028-2029`, `2029-2030`, `2030-2031`, `2031-2032` |

Các nhãn chỉ được thay trên workbook bản sao ở đúng 70 ô neo đã liệt kê;
không sửa mẫu gốc. Các năm lịch ở vùng ký xác nhận tại hàng 82 được giữ
nguyên. Cơ chế chỉ cập nhật nhãn, không ghi dữ liệu thành viên.

## 5. Bảng audit mapping B5–B12

“Ô đích” chỉ là vùng ứng viên theo nhãn/tiêu đề quan sát được. Đây chưa phải
mapping đã duyệt.

| Nhóm | Trường nguồn / cột | Ý nghĩa theo tên trường hiện tại | Ô đích ứng viên và nhãn năm | Quy tắc còn thiếu / bằng chứng | Câu hỏi người dùng cần quyết định |
|---|---|---|---|---|---|
| B5 | `grade_group` Q; `class_name` R | Khối và lớp hiện tại; hai trường không mang năm riêng | `H:J` trên một trong năm hàng có nhãn tại `F:G`; header “NHÓM, LỚP ĐANG HỌC (Theo năm học)” | Mẫu chứng minh vùng liên quan năm; không chứng minh dữ liệu hiện tại thuộc dòng nào. Nguồn Q/R không có trường năm kèm theo. | Q/R mô tả năm học nào? Dùng năm học cấu hình, dòng 1/2, hay giá trị khác? Định dạng kết hợp khối-lớp và cách xử lý R/ Q trống? |
| B6 | `school_district` S; `school_name` T | Huyện và tên trường hiện tại | `K:O` trên một trong năm hàng; header “TÊN TRƯỜNG - HUYỆN ĐANG HỌC (Tương ứng với từng năm học)” | Cấu trúc đích quan sát được; chưa có quy tắc ghép S/T, không có năm trên bản ghi nguồn. | Có ghép hai trường không; thứ tự/dấu phân cách; năm nào; xử lý một trường thiếu? |
| B7 | `graduation_level` V; `supplementary_education` W; `graduation_year` X | Bậc tốt nghiệp phổ thông, bổ túc, và năm tốt nghiệp | P (Cấp học), Q (Bổ túc), R (Năm học), trên một trong năm hàng | Tiêu đề đích quan sát được. X là trường có nghĩa “năm tốt nghiệp”, chưa xác minh đó là năm học hay năm dương lịch; chưa rõ quy tắc chọn hàng. | X lưu năm lịch hay năm học? Nếu năm lịch thì quy đổi thế nào? Khi nào ghi P/Q/R và xử lý hai loại cùng có dữ liệu? |
| B8 | `vocational_graduation_level` Y; `vocational_graduation_year` AA | Bậc tốt nghiệp nghề và năm tốt nghiệp nghề | S (Bậc), T:U (vùng ghi năm học theo mẫu), trên một trong năm hàng | Nhãn đích cho thấy vùng năm học; AA chưa được xác minh là năm học hay năm dương lịch. Chiều rộng merge cụ thể cần giữ nguyên. | AA mang loại năm nào? Dòng đích, định dạng, và cách ghi khi thiếu Y hoặc AA? |
| B9 | `completed_education` AB; `completed_education_year` AC | Trường trạng thái/mức học xong và năm hoàn thành theo tên cấu hình | V (Lớp), W (Năm học), dưới tiêu đề “Học xong” | Tọa độ ứng viên theo header; ngữ nghĩa/miền giá trị AB cần đối chiếu tiêu đề nguồn thực tế, năm AC chưa phân loại thành năm học hay năm lịch. | Giá trị AB biểu thị lớp, bậc hay trạng thái? AC là năm gì; quy tắc dòng và xử lý thiếu cặp? |
| B10 | `dropped_out_class` AD; `dropped_out_year` AE | Lớp bỏ học và năm bỏ học | X (Lớp), Y (Năm học), dưới tiêu đề “Bỏ học” | Header ứng viên rõ; AE chưa được xác minh là năm học hay năm dương lịch; chưa có quy tắc hàng. | AE là năm dương lịch hay năm học? Hàng đích nào, định dạng nào, và xử lý thiếu một trường? |
| B11 | `currently_studying_class` AF; `completed_class` AG | Lớp đang học và lớp đã hoàn thành | Z (Đang học lớp), AA (Hoàn thành lớp), các dòng năm trong khối | Nhãn đích quan sát được; source fields không có năm kèm theo. Chưa rõ xung đột với Q/R hoặc cách thể hiện đồng thời. | Dùng dòng năm học nào? Q/R có trùng AF không; nếu trùng hoặc mâu thuẫn, trường nào có thẩm quyền? |
| B12 | `reliteracy_level` AH | Mức tái mù chữ | AB (Tái mù chữ mức), các dòng năm trong khối | Header gần khớp theo tên; source AH không kèm năm, chưa rõ cần ghi theo năm hiện tại hay sự kiện. | Mức nguồn nào tương ứng từng giá trị đích? Dòng năm nào được dùng và khi nào để trống/cảnh báo? |

### Phân biệt các loại năm

- **Năm học hiện tại:** cấu hình `school_year`; chỉ có thể dùng làm ngữ
  cảnh của dữ liệu hiện tại sau khi người dùng duyệt điểm neo.
- **Năm học hiển thị:** nhãn năm trong `F:G`, lặp năm dòng mỗi thành viên.
- **Năm ghi nhận sự kiện:** chưa có trường nguồn độc lập được xác nhận cho
  mọi nhóm B5–B12.
- **Năm tốt nghiệp/bỏ học/hoàn thành:** X, AA, AC, AE có tên mô tả năm sự
  kiện. Không có bằng chứng hiện tại xác định chúng là năm dương lịch hay
  năm học; không tự đổi thành dòng.
- **Năm lịch ký xác nhận:** các nhãn “Năm 2025” đến “Năm 2029” ở hàng 82;
  tách riêng khỏi dãy năm học.

## 6. Kết quả thử nghiệm và kiểm thử

- Excel COM mở mẫu read-only thành công khi truy cập sheet theo chỉ số.
  Truy cập `Worksheets("Phiếu điều tra 1 mặt")` trực tiếp phát sinh
  `pywintypes.com_error`; đây là lỗi tra cứu tên sheet trong phiên COM,
  không phải lỗi workbook. Không ghi lên workbook.
- Đếm được 442 vùng merge, UsedRange `$A$1:$AG$94`, 70 nhãn năm học,
  A4 ngang, Zoom 68, hai ngắt ngang và một ngắt dọc như mô tả trên.
- Mẫu có `LinkSources(1) = None` khi đọc; không tạo workbook đầu ra nên
  kiểm tra lỗi Repair không áp dụng ở lượt này.
- Hash `maubaocao.xls` trước/sau: SHA-256
  `89354830a91694e6a9251ccd0b02214c19625c211da683e1a086a74411694aad`.
  Hash `dulieumau.xls` trước/sau: SHA-256
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
- Người dùng đã phê duyệt dòng nhãn thứ nhất làm điểm neo, giữ nguyên năm
  lịch ký và cho phép tạo bản sao thử nghiệm.
- Hai workbook thử nghiệm đã được tạo trong
  `output/testing/school_year_labels/`:
  `maubaocao_2026-2027.xlsx` và `maubaocao_2027-2028.xlsx`. Excel COM mở
  lại cả hai; mỗi workbook có đúng 70 nhãn theo dãy đã duyệt và 442 vùng
  gộp.
- Bổ sung kiểm thử `validate_export_settings`: chấp nhận `2026-2027` và
  `2027-2028`; từ chối sai định dạng và hai năm không liên tiếp. Kiểm thử
  này chỉ xác minh cấu hình, không phê duyệt quy tắc neo hoặc mapping.
- `.\venv\Scripts\python.exe -m unittest discover -v`: **28 PASS, 0 FAIL,
  0 SKIP**. Trong `test_school_year_labels.py`, 3 test unit kiểm tra tạo
  nhãn, từ chối cấu hình sai và allowlist; 2 test tích hợp Excel COM thử
  hai năm học trên bản sao và xác nhận hàm từ chối sai vùng gộp trước khi
  sửa giá trị. Bộ suite còn có kiểm thử Excel COM của exporter.
- Kiểm thử exporter xác nhận dữ liệu tổng hợp B5–B12 không bị ghi vào các
  ô ứng viên chưa được duyệt và hash nguồn/mẫu không đổi.
- `.\venv\Scripts\python.exe -m compileall -q .`: PASS.
- `config/settings.json` và `config/template_mapping.json`: parse JSON
  thành công; cấu hình giữ `school_year="2026-2027"` và `commune_name=null`.
- `git diff --check`: PASS.
- Kiểm tra đầy đủ so với bản gốc trên các bản copy: danh sách ô có giá trị
  thay đổi đúng 70 nhãn; giá trị các ô còn lại, vùng ký, merge, font/viền
  của nhãn, chiều cao dòng, chiều rộng cột, PrintArea, ngắt trang, thiết
  lập trang và danh sách external link không đổi. Không có hộ thật hoặc dữ
  liệu thành viên được ghi vào hai workbook thử.
