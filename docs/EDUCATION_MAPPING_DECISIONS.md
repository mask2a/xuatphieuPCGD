# Phân tích dữ liệu nguồn cho mapping B5–B12

Ngày kiểm tra: 2026-10-20
Nguồn: `dulieumau.xls`, sheet `MauNhapLieu`, mở bằng Excel COM với
`ReadOnly=True`, `UpdateLinks=0`.
Mẫu: `maubaocao.xls`, mở bằng Excel COM với `ReadOnly=True`.

## Phạm vi và cách đọc

Đọc các bản ghi có dữ liệu trong vùng A:AH, bắt đầu tại dòng 5. Có 427 bản
ghi thực, từ dòng 5 đến dòng 431. Các hàng định dạng trống phía dưới không
được tính là bản ghi. Thống kê không gồm họ tên, số điện thoại, mã hộ hay
thông tin định danh cá nhân. Không chuẩn hóa, sửa hoặc ghi vào nguồn, mẫu hay
`config/template_mapping.json`; không tạo workbook đầu ra.

Giá trị số được phân loại theo `Value2` của Excel COM; chuỗi giữ nguyên nội
dung thấy trong ô. Không có giá trị ngày Excel trong các cột được khảo sát.
Các giá trị số ít gặp ở Q/R vẫn được giữ riêng, không ép thành chuỗi hay sửa
trong dữ liệu.

## 1. Thống kê theo cột

| Cột | Tiêu đề nguồn | Có dữ liệu | Trống | Giá trị khác nhau* | Kiểu dữ liệu quan sát |
|---|---|---:|---:|---:|---|
| Q | Khối học | 105 | 322 | 17 | 104 chuỗi, 1 số |
| R | Lớp học | 113 | 314 | 64 | 111 chuỗi, 2 số |
| S | Quận/Huyện đang học | 1 | 426 | 1 | 1 chuỗi |
| T | Tên trường đang học | 113 | 314 | 19 | 113 chuỗi |
| V | Bậc tốt nghiệp | 304 | 123 | 4 | 304 chuỗi |
| W | Bổ túc | 0 | 427 | 0 | Không có dữ liệu |
| X | Năm tốt nghiệp | 301 | 126 | 49 | 301 chuỗi |
| Y | Bậc TN Nghề | 0 | 427 | 0 | Không có dữ liệu |
| AA | Năm TN nghề | 0 | 427 | 0 | Không có dữ liệu |
| AB | Học xong | 0 | 427 | 0 | Không có dữ liệu |
| AC | Năm (của Học xong) | 0 | 427 | 0 | Không có dữ liệu |
| AD | Bỏ học – lớp | 2 | 425 | 1 | 2 chuỗi |
| AE | Bỏ học – năm | 2 | 425 | 1 | 2 chuỗi |
| AF | Đang học lớp | 0 | 427 | 0 | Không có dữ liệu |
| AG | Hoàn thành lớp | 0 | 427 | 0 | Không có dữ liệu |
| AH | Tái mù chữ mức | 0 | 427 | 0 | Không có dữ liệu |

\* Giá trị khác nhau được đếm theo cả nội dung và kiểu: ví dụ số `11` và
chuỗi `"11"` ở Q là hai giá trị riêng.

## 2. Giá trị thực tế khác nhau và số lần xuất hiện

### Q — Khối học

Chuỗi: `"1" ×3`; `"2" ×8`; `"3" ×5`; `"4" ×9`; `"5" ×8`; `"6" ×5`;
`"7" ×7`; `"8" ×12`; `"9" ×14`; `"10" ×5`; `"11" ×9`; `"12" ×8`;
`"13-24 tháng" ×3`; `"3-4 tuổi" ×2`; `"4-5 tuổi" ×5`; `"TDN1" ×1`.
Ngoài ra có giá trị **số** `11 ×1`.

### R — Lớp học

Giá trị số: `6 ×2`.

Chuỗi: `"1A1*" ×2`; `"1A2*" ×1`; `"2A1*" ×2`; `"2A2*" ×3`; `"2A4*" ×2`;
`"2A8*" ×1`; `"3A1*" ×2`; `"3A2*" ×1`; `"3A3*" ×2`; `"4A1*" ×3`;
`"4A2*" ×2`; `"4A3*" ×3`; `"4A4*" ×1`; `"5A2*" ×1`; `"5A3*" ×2`;
`"5A4*" ×3`; `"5A5*" ×1`; `"5A7*" ×1`; `"6A10" ×1`; `"6A3" ×3`;
`"6A4" ×4`; `"6A6" ×1`; `"6A7" ×1`; `"6a4" ×1`; `"7A3" ×1`;
`"7A4" ×3`; `"7A5" ×2`; `"7A7" ×1`; `"8" ×1`; `"8A12" ×1`;
`"8A3" ×1`; `"8A4" ×4`; `"8A5" ×3`; `"8A6" ×1`; `"9A3" ×4`;
`"9A4" ×2`; `"9A5" ×3`; `"9A6" ×2`; `"9A8" ×3`; `"10" ×1`;
`"10A1" ×1`; `"10A6" ×1`; `"10A9" ×1`; `"11A1" ×1`; `"11A13" ×1`;
`"11A2" ×2`; `"11A3" ×1`; `"12" ×1`; `"12A1" ×3`; `"12A2" ×2`;
`"12A5" ×1`; `"12A6" ×1`; `"11 CHUYÊN NGA" ×1`;
`"11 CHUYÊN TRUNG" ×1`; `"11 chuyên Sinh" ×1`; `"CHUYỂN" ×1`;
`"CHUYỂN BN" ×1`; `"CN Ô TÔ" ×1`; `"M2*" ×3`; `"M3**" ×2`;
`"M4**" ×5`; `"Nghề" ×1`; `"THPT" ×1`.

### S — Quận/Huyện đang học

`"MN Trưng Vương" ×1`.

### T — Tên trường đang học

`"1982-1983" ×1`; `"Cao đẳng Than khoáng sản" ×3`; `"CĐNCNVB" ×1`;
`"MN Hoa Hướng Dương" ×1`; `"MN Khánh Hòa" ×1`; `"MN Tân Lập" ×1`;
`"MN Tư thục Hoa Tuổi Thơ" ×1`; `"Mầm non Sơn Cẩm" ×6`;
`"THCS Quan Triều 2" ×35`; `"THCS Sơn Cẩm 1" ×1`;
`"THCS Trưng Vương" ×2`; `"THPT Chuyên" ×3`; `"THPT Khánh Hoà" ×1`;
`"THPT Khánh Hòa" ×19`; `"THPT Lương Ngọc Quyến" ×3`;
`"TTGDTT tỉnh Thái Nguyên" ×1`; `"Tiểu học Quan Triều 2" ×29`;
`"Tiểu học Đại Phúc" ×2`; `"Tiểu học Đội Cấn" ×2`.

### V — Bậc tốt nghiệp

`"MN" ×33`; `"TH" ×46`; `"THCS" ×35`; `"THPT" ×190`.

### X — Năm tốt nghiệp

Tất cả là chuỗi. Mọi giá trị đều có dạng `YYYY-YYYY` với hai năm liên tiếp.
Giá trị/count:

`1960-1961 ×1`; `1962-1963 ×1`; `1965-1966 ×2`; `1967-1968 ×1`;
`1969-1970 ×1`; `1971-1972 ×1`; `1981-1982 ×2`; `1982-1983 ×4`;
`1983-1984 ×1`; `1984-1985 ×4`; `1986-1987 ×7`; `1987-1988 ×4`;
`1988-1989 ×1`; `1989-1990 ×7`; `1990-1991 ×1`; `1991-1992 ×4`;
`1992-1993 ×3`; `1993-1994 ×7`; `1994-1995 ×4`; `1995-1996 ×1`;
`1996-1997 ×6`; `1997-1998 ×4`; `1998-1999 ×5`; `1999-2000 ×6`;
`2000-2001 ×3`; `2001-2002 ×4`; `2002-2003 ×5`; `2003-2004 ×5`;
`2004-2005 ×11`; `2005-2006 ×4`; `2006-2007 ×6`; `2007-2008 ×6`;
`2008-2009 ×3`; `2009-2010 ×1`; `2010-2011 ×3`; `2011-2012 ×6`;
`2012-2013 ×4`; `2014-2015 ×5`; `2015-2016 ×3`; `2016-2017 ×1`;
`2017-2018 ×2`; `2018-2019 ×15`; `2019-2020 ×4`; `2020-2021 ×7`;
`2021-2022 ×15`; `2022-2023 ×27`; `2023-2024 ×34`;
`2024-2025 ×30`; `2025-2026 ×19`.

### AD, AE — Bỏ học

- AD: chuỗi `"10" ×2`.
- AE: chuỗi `"1988-1989" ×2`.

### Các cột không có giá trị khác nhau

W, Y, AA, AB, AC, AF, AG và AH đều trống trong cả 427 bản ghi; vì vậy
không thể kiểm tra từ điển giá trị hay quy tắc tương ứng dựa trên tệp này.

## 3. Mối liên hệ và điểm cần lưu ý

- **Q/R:** có 102 bản ghi điền cả hai cột; 3 chỉ có Q và 11 chỉ có R.
  Trong 102 cặp đầy đủ, 90 cặp có thể so sánh phần số đầu lớp R với Q và
  phần số đó trùng nhau; 12 cặp không có dạng số để đối chiếu theo cách
  này. Đây chỉ là quan sát cấu trúc, không xác nhận Q/R là cùng một nguồn
  thẩm quyền hay thuộc `school_year` hiện tại.
- Q/R có dữ liệu trên tổng cộng 116 bản ghi. Q có ba nhóm tuổi mầm non,
  `TDN1`, các giá trị chữ/số và một giá trị số; R có ký hiệu lớp khác kiểu
  chữ hoa/thường, hai giá trị số, và một số giá trị văn bản như `CHUYỂN`,
  `CN Ô TÔ`, `Nghề`, `THPT`. Đây là các giá trị cần giữ nguyên và cần quy
  tắc xử lý, không tự sửa thành lớp chuẩn.
- **S/T:** có dữ liệu trên 114 bản ghi nhưng không có bản ghi nào đồng
  thời có cả S và T: 1 bản ghi chỉ có S, 113 chỉ có T. Giá trị duy nhất của
  S có hình thức giống tên cơ sở hơn là tên huyện; T có một giá trị
  `1982-1983` giống chuỗi năm. Đây là dấu hiệu cần kiểm tra căn chỉnh/ý
  nghĩa cột, không đủ căn cứ kết luận ô nào sai hoặc chuyển dữ liệu.
- **V/W/X:** có dữ liệu trên 305 bản ghi; V và X cùng có dữ liệu ở 300,
  V chỉ có ở 4, X chỉ có ở 1; W trống cả 427. Bậc V có bốn chuỗi MN, TH,
  THCS, THPT. X có 301 giá trị thuộc 49 năm dạng `YYYY-YYYY`, từ
  1960-1961 đến 2025-2026.
- **Năm sự kiện so với năm cấu hình:** cấu hình hiện tại là
  `2026-2027`. Cả 301 giá trị X và 2 giá trị AE đều là cặp năm liên tiếp
  có năm bắt đầu trước 2026; không giá trị nào bằng `2026-2027`. Dạng
  `YYYY-YYYY` là bằng chứng về hình thức dữ liệu, nhưng tên cột chỉ ghi
  “Năm”; chưa đủ căn cứ khẳng định nghiệp vụ là năm học hay để xác định
  một trong năm dòng hiện hành trên phiếu.
- **Y/AA, AB/AC, AF/AG, AH:** hoàn toàn trống; không thể kiểm tra mối
  liên hệ giữa bậc/năm tốt nghiệp nghề, học xong/năm, đang học/lớp hoàn
  thành hoặc mức tái mù chữ từ dữ liệu hiện tại.
- **AD/AE:** cả hai bản ghi đều có đủ cặp lớp-năm; năm có dạng
  `1988-1989`. Cỡ mẫu 2 không xác định được quy tắc hàng đích hoặc cách
  nối năm sự kiện với nhãn năm của phiếu.

## 4. Ô đích ứng viên quan sát trực tiếp trên mẫu

Excel COM read-only xác nhận tiêu đề và các vùng gộp sau. Đây là vị trí
ứng viên theo tiêu đề mẫu, **không phải mapping đã phê duyệt**:

| Nhóm | Nguồn | Tiêu đề/vùng đích ứng viên trên mỗi dòng thành viên |
|---|---|---|
| B5 | Q, R | H:J — “NHÓM, LỚP ĐANG HỌC (Theo năm học)” |
| B6 | S, T | K:O — “TÊN TRƯỜNG - HUYỆN ĐANG HỌC (Tương ứng với từng năm học)” |
| B7 | V, W, X | P — “Cấp học”; Q — “Bổ túc”; R — “Năm học” |
| B8 | Y, AA | S — “Bậc”; T:U — “Năm học” |
| B9 | AB, AC | V — “Lớp”; W — “Năm học” |
| B10 | AD, AE | X — “Lớp”; Y — “Năm học” |
| B11 | AF, AG | Z — “Đang học lớp”; AA — “Hoàn thành lớp” |
| B12 | AH | AB — “Tái mù chữ mức” |

Các vùng ứng viên thuộc từng dòng thành viên. Các cột năm của phiếu mang
nhãn 5 năm do cấu hình `school_year` điều khiển; việc đã duyệt cập nhật
nhãn không xác định dòng nào phải nhận một sự kiện trong nguồn.

## 5. Đề xuất dựa trên bằng chứng — chưa ghi vào mapping

1. Giữ nguyên giá trị nguồn, kiểu dữ liệu và chuỗi năm; không ghép/chuẩn
   hóa Q/R, S/T hoặc đổi cặp năm thành năm đơn.
2. Không tự đặt các sự kiện tốt nghiệp/bỏ học lịch sử vào một trong năm
   dòng hiện hành. X và AE có năm sớm hơn các nhãn hiện hành, nên cần quy
   tắc riêng về nơi thể hiện trước khi ghi.
3. Với S/T, tạm không ghép hai cột vì không có bản ghi đồng thời có cả
   hai. Cần xác minh ý nghĩa và căn chỉnh nguồn trước khi quyết định cách
   in trường/huyện.
4. Các cột hoàn toàn trống không cung cấp bằng chứng để quyết định mapping
   hoặc từ điển giá trị. Có thể chỉ định chưa ánh xạ/để trống và cảnh báo,
   nhưng cần người dùng phê duyệt hành vi này.
5. Chưa thay đổi `config/template_mapping.json`; không có quy tắc nghiệp vụ
   B5–B12 nào được tự phê duyệt trong báo cáo này.

## 6. Câu hỏi ngắn cần người dùng phê duyệt

1. **B5:** Q/R có đại diện cho đúng `school_year` cấu hình `2026-2027` và
   được ghi vào dòng nhãn đầu tiên không? Nếu Q/R khác nhau, một trường
   thiếu, hoặc là nhóm tuổi/`TDN1`/văn bản thay vì khối-lớp phổ thông thì
   giữ nguyên, để trống hay cảnh báo?
2. **B6:** Có thể xác nhận S là huyện và T là trường theo từng bản ghi
   không? Hiện hai cột không giao nhau; xử lý thế nào với S đơn lẻ, T đơn
   lẻ và giá trị năm dạng chuỗi trong T? Chưa ghép hai trường cho đến khi
   được xác nhận.
3. **B7:** X có dạng năm học liên tiếp, nhưng mọi giá trị đều trước
   `2026-2027`. Có coi X là năm học không? Sự kiện tốt nghiệp lịch sử phải
   ghi ở đâu khi không trùng năm dòng đang hiển thị? Xác nhận cách ghi V
   vào P, W vào Q (W hiện trống), và xử lý bốn V thiếu X cùng một X thiếu V.
4. **B8:** Y/AA đều trống. Có chấp thuận giữ mapping chưa hoạt động/để
   trống và cảnh báo khi sau này có dữ liệu, hay cần cung cấp quy tắc từ
   nguồn nghiệp vụ khác?
5. **B9:** AB/AC đều trống. Có chấp thuận để trống và cảnh báo khi phát
   sinh dữ liệu, hay cung cấp quy tắc và ý nghĩa trường từ tài liệu khác?
6. **B10:** Hai cặp AD/AE đều đầy đủ; AE là `1988-1989`. Có xác nhận AE là
   năm học không, và sự kiện bỏ học lịch sử sẽ ghi ở dòng nào khi phiếu
   hiện chỉ hiển thị các năm từ cấu hình trở đi?
7. **B11:** AF/AG đều trống. Khi có dữ liệu, có xác nhận đó là thông tin
   của năm `school_year` hiện hành và nằm ở dòng nhãn đầu không? Nếu mâu
   thuẫn với Q/R thì trường nào có thẩm quyền?
8. **B12:** AH trống. Có chấp thuận chưa ánh xạ và cảnh báo giá trị chưa
   được xác nhận khi phát sinh, hay cung cấp bảng quy đổi mức tái mù chữ
   cùng năm/dòng áp dụng?

## 7. Kiểm tra an toàn

- Excel COM đã mở `dulieumau.xls` và `maubaocao.xls` ở chế độ chỉ đọc.
- SHA-256 sau khi phân tích: nguồn
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`;
  mẫu `89354830a91694e6a9251ccd0b02214c19625c211da683e1a086a74411694aad`.
- Không sửa nguồn/mẫu, không tạo workbook kết quả, không sửa mapping, không
  xuất hộ thật và không commit.
- Không chạy unittest vì không thay đổi mã nguồn. Giai đoạn 3 vẫn chưa
  nghiệm thu; các quyết định B5–B12 đang chờ người dùng phê duyệt.

## Phê duyệt và thử mapping B5–B7 — 2026-10-20

Sau khi lập phân tích này, người dùng đã duyệt triển khai thử:

| Nhóm | Nguồn ghi | Ô neo đã xác minh trên dòng đầu | Quy tắc |
|---|---|---|---|
| B5 | R (`class_name`); Q chỉ để đối chiếu | H tại các hàng 9, 14, 19, 24, 29, 34, 39, 47, 52, 57, 62, 67, 72, 77 | Không ghép Q/R; giữ giá trị R kể cả giá trị đặc biệt; cảnh báo mâu thuẫn/không so sánh được |
| B6 | T (`school_name`) | K tại cùng 14 hàng | S chưa duyệt; T có dạng năm phải cảnh báo và giữ nguyên |
| B7 | V và X | P/R tại cùng 14 hàng | Ghi nguyên dữ liệu ở hàng đầu, kể cả năm quá khứ; W không ghi; thiếu V hoặc X thì cảnh báo |

Excel COM read-only đã xác minh cho cả 14 vị trí: H là neo vùng gộp H:J,
K là neo K:O; P và R không gộp. Các địa chỉ và merge-area được khai báo
tường minh trong `config/template_mapping.json`, và được kiểm tra khi tải
cấu hình cũng như trước khi ghi.

Lựa chọn riêng của người dùng cho lần thử là giữ nguyên chính xác 70 nhãn
năm có trên mẫu. Vì nhãn đầu trên mẫu là `2025-2026`, còn cấu hình là
`2026-2027`, các workbook thử dưới đây không đồng bộ nhãn theo cấu hình;
việc đặt B5–B7 ở hàng đầu chỉ kiểm chứng vị trí, không xác nhận năm học
hiển thị.

Ba workbook tổng hợp dùng dữ liệu giả lập, không chứa dữ liệu hộ thật:

- `output/testing/education_mapping/ho_SYNTHETIC-06.xlsx`
- `output/testing/education_mapping/ho_SYNTHETIC-07.xlsx`
- `output/testing/education_mapping/ho_SYNTHETIC-14.xlsx`

Kết quả kiểm thử kỹ thuật, workbook mở lại bằng Excel COM: các giá trị R/T/
V/X được ghi nguyên tại H/K/P/R; sự kiện tốt nghiệp cũ vẫn giữ chuỗi năm;
M4**, CHUYỂN và tên trường dạng chuỗi năm được giữ nguyên kèm cảnh báo;
Q/R không khớp và V/X thiếu một trường có cảnh báo. Các dòng năm thứ 2–5
không nhận dữ liệu B5–B7; dữ liệu không bị chuyển sang thành viên kế tiếp;
70 nhãn giữ nguyên như mẫu, 442 vùng gộp và định dạng ô đích được kiểm tra.
Mẫu gốc không đổi.

Đây là kiểm thử kỹ thuật trên dữ liệu tổng hợp, không phải nghiệm thu
nghiệp vụ. B8–B12, S, W và các mapping khác vẫn chưa được duyệt; không
xuất dữ liệu thật.
