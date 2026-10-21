# Quyết định nghiệp vụ cần người dùng phê duyệt

Ngày lập: 2026-10-20

Tài liệu này trình bày 15 nhóm B1–B15 từ
[MAPPING_AUDIT.md](./MAPPING_AUDIT.md). Các ô đích dưới đây chỉ là vị trí
ứng viên đã quan sát trên mẫu, không phải mapping được phép dùng. Không có
quyết định nào được tự chọn hoặc ghi vào cấu hình ánh xạ trong tài liệu này.

## Tình trạng chờ phê duyệt

- Tệp nguồn hiện tại `dulieumau.xls` có SHA-256:
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
- Digest này khác với giá trị trong báo cáo đối chiếu cũ. Chưa có bản nguồn
  cũ và chưa xác định được nguyên nhân khác biệt.
- Trạng thái phiên bản nguồn: **CHỜ NGƯỜI DÙNG PHÊ DUYỆT**. Không xuất hộ
  thật cho đến khi người dùng xác nhận phiên bản nguồn được phép sử dụng.
- Cấu hình năm học là `2026-2027`; `commune_name` hiện là `null`. Tên
  xã/phường phải do người dùng cung cấp; chưa thể xuất chính thức khi giá
  trị này chưa được xác nhận.
- Người dùng đã chọn **Phương án A – Năm học linh hoạt**. Quyết định này
  cho phép thiết kế cập nhật nhãn năm học trên bản sao, nhưng chưa duyệt
  điểm neo của năm cấu hình và chưa duyệt ánh xạ dữ liệu thành viên B5–B12.

## Nhóm quyết định

Với các nhóm B1–B14, quyết định có thể ảnh hưởng trực tiếp tới dữ liệu in
nếu trường nguồn có giá trị. Có thể hoãn một nhóm chỉ khi người dùng chấp
thuận rõ việc để trống và tiếp tục cảnh báo; hoãn không đồng nghĩa với
mapping đã hoàn tất. B15 được đề xuất là tùy chọn, nhưng nếu cột nguồn có
ghi chú thì lựa chọn vẫn ảnh hưởng nội dung in. Tính bắt buộc pháp lý hay
nghiệp vụ của bất kỳ trường nào chưa được xác nhận.

## Xác nhận bổ sung cho Phương án A — điểm neo của năm học

Người dùng đã duyệt:

1. `school_year` là nhãn đầu tiên trong mỗi khối thành viên. Năm bắt đầu
   Y và chỉ số dòng i từ 0 đến 4 tạo nhãn `Y+i` đến `Y+i+1`.
2. Năm lịch trong phần ký xác nhận phải giữ nguyên.
3. Được phép tạo bản sao thử nghiệm chỉ đổi nhãn năm học, rồi kiểm tra lại
   bằng Excel COM.

Kết quả quy tắc đã kiểm thử:

- `2026-2027` → `2026-2027`, `2027-2028`, `2028-2029`, `2029-2030`,
  `2030-2031`.
- `2027-2028` → `2027-2028`, `2028-2029`, `2029-2030`, `2030-2031`,
  `2031-2032`.

Danh sách 70 ô neo và kết quả so sánh hai workbook thử được ghi trong
[SCHOOL_YEAR_MAPPING_AUDIT.md](./SCHOOL_YEAR_MAPPING_AUDIT.md). Quyết định
điểm neo không phê duyệt mapping nghiệp vụ B5–B12.

Phê duyệt Phương án A **không đồng nghĩa** phê duyệt ánh xạ B5–B12. Bảng
phân tích từng nhóm, trường/cột và câu hỏi theo loại năm nằm trong
[SCHOOL_YEAR_MAPPING_AUDIT.md](./SCHOOL_YEAR_MAPPING_AUDIT.md).

## Phê duyệt thử mapping B5–B7 — 2026-10-20

Người dùng phê duyệt triển khai thử theo các quy tắc sau:

- **B5:** Lấy lớp cụ thể ở R để ghi vào ô H của dòng nhãn năm thứ nhất
  trong khối thành viên. Q chỉ dùng đối chiếu; không ghép Q/R và không tự
  sửa giá trị đặc biệt như `M4**`, `CHUYỂN` hoặc `Nghề`. Giá trị bất thường,
  không so sánh được hoặc mâu thuẫn phải cảnh báo.
- **B6:** Lấy tên trường ở T để ghi vào ô K của dòng thứ nhất. S chưa được
  duyệt ánh xạ. T có dạng bất thường như chuỗi năm thì ghi cảnh báo nhưng
  giữ nguyên giá trị.
- **B7:** Ghi V vào P và X vào R của dòng thứ nhất, kể cả năm tốt nghiệp
  trước năm cấu hình. Giữ nguyên chuỗi năm. W chưa được duyệt; nếu V hoặc X
  thiếu thì ghi phần có dữ liệu và cảnh báo.
- B10 và các mapping khác không được duyệt thì không ghi.
- Cho bản sao thử giữ nguyên toàn bộ 70 nhãn năm hiện có của mẫu. Hiện mẫu
  bắt đầu ở `2025-2026`, khác cấu hình `2026-2027`; lựa chọn giữ nhãn này
  chỉ áp dụng cho lượt thử, không xác nhận rằng cấu hình và nhãn đã đồng bộ.

### B1 — Giới tính

- **Trường nguồn:** `female`, cột G. Giá trị `X` được mô tả là nữ.
- **Ô đích dự kiến:** `C{s+3}` trong khối thành viên, cạnh nhãn “Nữ:”.
  Mẫu có ký tự `q` font Wingdings; ký tự/hiển thị để chọn chưa được xác minh.
- **Cần người dùng xác nhận:** Cung cấp phiếu đã điền chuẩn hoặc quy ước
  chính thức về ký tự đánh dấu. Xác nhận thêm cách xử lý G trống hoặc khác
  `X`.
- **Các lựa chọn cần duyệt:** Dùng đúng dấu theo phiếu/hướng dẫn đã duyệt;
  hoặc để ô trống và giữ cảnh báo cho tới khi có quy ước. Không tự suy ra
  cách ghi giới tính còn lại.
- **Ảnh hưởng:** Trực tiếp tới nội dung giới tính trên phiếu; không nên
  triển khai trước khi dấu được xác minh.

### B2 — Diện cư trú

- **Trường nguồn:** `residence`, cột O.
- **Ô đích dự kiến:** `Z2:AG2`, hai lựa chọn in trên mẫu là “Thường trú” và
  “Tạm trú”.
- **Cần người dùng xác nhận:** Chọn dấu/placeholder chính xác cho từng diện;
  xác nhận chuẩn hóa khác biệt hoa thường; quyết định xử lý khi thành viên
  trong cùng hộ có diện cư trú khác nhau.
- **Các lựa chọn cần duyệt:** Chọn một trong hai diện theo dữ liệu nhất
  quán; hoặc khi mâu thuẫn để trống và cảnh báo/chờ duyệt thủ công. Chỉ chọn
  theo hồ sơ khi có quy trình và căn cứ được người dùng cung cấp.
- **Ảnh hưởng:** Trực tiếp tới thông tin hộ và quyết định xuất phiếu.

### B3 — Địa chỉ

- **Trường nguồn:** `address`, cột M.
- **Ô đích dự kiến:** `A3:G3`, nhãn “Xóm (2)”.
- **Cần người dùng xác nhận:** Cột M có được chép nguyên văn hay chỉ lấy
  thành phần xóm/tổ; cách xử lý địa chỉ dài và địa chỉ không có thành phần
  xóm.
- **Các lựa chọn cần duyệt:** In nguyên văn M; chỉ in thành phần xóm theo
  quy tắc tách dữ liệu người dùng cung cấp; hoặc để trống/cảnh báo khi
  không tách được.
- **Ảnh hưởng:** Trực tiếp tới địa chỉ in trên phiếu.

### B4 — Tình trạng cư trú cá nhân

- **Trường nguồn:** `residence_status`, cột P.
- **Ô đích dự kiến:** Cột AF trong khối thành viên; mẫu có nhãn chuyển đến,
  chuyển đi, chết.
- **Cần người dùng xác nhận:** Danh sách giá trị P có thể có và ý nghĩa;
  giá trị nào tương ứng từng nhãn đích; cách xử lý giá trị không khớp.
- **Các lựa chọn cần duyệt:** Cung cấp bảng quy đổi chính thức; hoặc giữ
  trống và cảnh báo các giá trị chưa có quy tắc.
- **Ảnh hưởng:** Trực tiếp tới trạng thái của từng thành viên nếu nguồn có
  dữ liệu.

### B5 — Khối/lớp hiện tại

- **Trạng thái:** Đã duyệt để triển khai thử; chưa nghiệm thu nghiệp vụ.
- **Trường nguồn:** `class_name` (R) được ghi; `grade_group` (Q) chỉ dùng
  để đối chiếu.
- **Ô đích:** Ô neo H tại dòng thứ nhất của từng khối thành viên (các địa
  chỉ theo từng vị trí được liệt kê tường minh trong
  `config/template_mapping.json`).
- **Quy tắc:** Không nối Q và R; không tự chuyển đổi/sửa `M4**`, `CHUYỂN`,
  `Nghề` hoặc giá trị đặc biệt khác. Cảnh báo khi Q/R mâu thuẫn hoặc không
  đủ dạng để so sánh; vẫn giữ nguyên lớp ở R. Không dùng Q thay thế khi R
  trống.
- **Giới hạn:** Không ánh xạ `currently_studying_class` (AF). Trong workbook
  thử lần này, nhãn năm giữ nguyên mẫu, không cập nhật từ `school_year`.

### B6 — Huyện và trường đang học

- **Trạng thái:** Đã duyệt ghi riêng tên trường; ánh xạ huyện chưa được duyệt.
- **Trường nguồn:** `school_name` (T) được ghi; `school_district` (S) không
  ghi.
- **Ô đích:** Ô neo K tại dòng thứ nhất của từng khối thành viên.
- **Quy tắc:** Ghi nguyên giá trị T. Nếu T có dạng chuỗi năm hoặc giá trị
  bất thường thì cảnh báo, không sửa và không đưa S vào cùng ô.

### B7 — Tốt nghiệp phổ thông/bổ túc

- **Trạng thái:** Đã duyệt V/X để triển khai thử; W vẫn chưa được duyệt.
- **Trường nguồn:** `graduation_level` (V) và `graduation_year` (X) được
  ghi riêng; `supplementary_education` (W) không ghi.
- **Ô đích:** P (bậc) và R (năm) ở dòng thứ nhất của từng khối thành viên.
- **Quy tắc:** Ghi nguyên chuỗi năm X kể cả sự kiện trước năm cấu hình;
  không đổi năm, đổi dòng hay suy diễn từ nhãn năm. Nếu V hoặc X thiếu,
  ghi thành phần còn có và phát cảnh báo.

### B8 — Tốt nghiệp nghề

- **Trường nguồn:** `vocational_graduation_level` (Y),
  `vocational_graduation_year` (AA).
- **Ô đích dự kiến:** Cột S (bậc) và vùng T:U (năm học) trong khối thành
  viên.
- **Cần người dùng xác nhận:** Dòng/năm áp dụng, định dạng năm, cách ghi
  khi thiếu một trong hai trường. AA là năm dương lịch hay năm học; không
  đổi thành dòng năm học nếu chưa có bảng quy đổi.
- **Các lựa chọn cần duyệt:** Ghi bậc và năm theo quy tắc được duyệt; hoặc
  để trống/cảnh báo nếu thiếu quy tắc hay dữ liệu ghép cặp.
- **Ảnh hưởng:** Trực tiếp tới thông tin tốt nghiệp nghề nếu có dữ liệu.

### B9 — Học xong

- **Trường nguồn:** `completed_education` (AB),
  `completed_education_year` (AC).
- **Ô đích dự kiến:** Cột V (trạng thái) và W (năm) trong khối thành viên.
- **Cần người dùng xác nhận:** Dòng năm cần ghi; quy tắc liên hệ giữa trạng
  thái và năm; AC là năm dương lịch hay năm học; cách xử lý khi chỉ một
  trường có giá trị.
- **Các lựa chọn cần duyệt:** Ghi theo cặp trên dòng/năm người dùng chỉ
  định; hoặc giữ trống/cảnh báo khi thiếu quy tắc hoặc thiếu thành phần.
- **Ảnh hưởng:** Trực tiếp tới tình trạng hoàn thành học tập nếu có dữ liệu.

### B10 — Bỏ học

- **Trường nguồn:** `dropped_out_class` (AD), `dropped_out_year` (AE).
- **Ô đích dự kiến:** Cột X (lớp) và Y (năm) trong khối thành viên.
- **Cần người dùng xác nhận:** Ý nghĩa/định dạng lớp và năm, dòng đích, và
  cách xử lý khi chỉ một trường có dữ liệu. AE là năm dương lịch hay năm
  học; nếu là năm dương lịch, người dùng cần cung cấp quy tắc đổi sang dòng
  năm học (nếu có).
- **Các lựa chọn cần duyệt:** Ghi theo cặp lớp-năm ở dòng được chỉ định;
  hoặc để trống/cảnh báo khi chưa xác định được.
- **Ảnh hưởng:** Trực tiếp tới thông tin bỏ học nếu có dữ liệu.

### B11 — Đang học và lớp hoàn thành

- **Trường nguồn:** `currently_studying_class` (AF),
  `completed_class` (AG).
- **Ô đích dự kiến:** Cột Z (đang học) và AA (hoàn thành) trong khối thành
  viên.
- **Cần người dùng xác nhận:** Dòng năm áp dụng; cách phân biệt hai trạng
  thái; cách ghi nếu cả hai cùng có dữ liệu hoặc mâu thuẫn với Q/R. AF/AG
  không kèm năm riêng; xác nhận đây là thông tin của `school_year` hiện
  hành hay một mốc khác.
- **Các lựa chọn cần duyệt:** Ghi từng trạng thái vào cột đích tương ứng
  theo năm được chỉ định; hoặc giữ trống/cảnh báo khi chưa có quy tắc về
  ưu tiên/xung đột.
- **Ảnh hưởng:** Trực tiếp tới tình trạng học tập nếu có dữ liệu.

### B12 — Tái mù chữ

- **Trường nguồn:** `reliteracy_level`, cột AH.
- **Ô đích dự kiến:** Cột AB trong khối thành viên.
- **Cần người dùng xác nhận:** Từ điển giá trị nguồn, ý nghĩa của từng mức
  và dòng năm đích. AH không kèm năm; xác nhận mức tái mù chữ thuộc năm
  học hiện hành hay cần ghi ở một năm/sự kiện khác.
- **Các lựa chọn cần duyệt:** Cung cấp bảng đối chiếu mức; hoặc để
  trống/cảnh báo giá trị chưa được giải thích.
- **Ảnh hưởng:** Trực tiếp nếu hộ có trường hợp tái mù chữ.

### B13 — Các loại khuyết tật

- **Trường nguồn:** `disability_mobility`, `disability_hearing_speech`,
  `disability_vision`, `disability_neurological`,
  `disability_intellectual`, `disability_learning`, `disability_autism`,
  `disability_other`, các cột AI:AP.
- **Ô đích dự kiến:** Cột AC (“Loại”) trong khối thành viên.
- **Cần người dùng xác nhận:** Giá trị nào trong mỗi cột có nghĩa là có
  khuyết tật; cách ghi nhiều loại vào một ô; thứ tự, dấu phân cách, cách ghi
  loại “khác” và dòng năm.
- **Các lựa chọn cần duyệt:** Ghi một hoặc nhiều loại theo danh sách/thứ
  tự/dấu phân cách được duyệt; hoặc để trống/cảnh báo cho đến khi quy ước
  được xác nhận.
- **Ảnh hưởng:** Trực tiếp nếu nguồn ghi nhận khuyết tật.

### B14 — Chứng nhận và khả năng học tập

- **Trường nguồn:** `disability_certified` (AQ),
  `learning_capacity` (AR).
- **Ô đích dự kiến:** Cột AD và AE trong khối thành viên.
- **Cần người dùng xác nhận:** Từ điển giá trị, dòng/năm cần ghi và cách
  liên hệ với loại khuyết tật ở B13.
- **Các lựa chọn cần duyệt:** Cung cấp bảng quy đổi được duyệt; hoặc để
  trống/cảnh báo giá trị chưa được xác định.
- **Ảnh hưởng:** Trực tiếp nếu nguồn có thông tin chứng nhận/khả năng học.

### B15 — Ghi chú thành viên

- **Trường nguồn:** `notes`, cột AX.
- **Ô đích dự kiến:** Cột AG trong khối thành viên.
- **Cần người dùng xác nhận:** Dòng nào trong 5 dòng của thành viên dùng
  cho ghi chú; có cho xuống dòng không; giới hạn độ dài và cách xử lý nội
  dung dài.
- **Các lựa chọn cần duyệt:** Ghi tại dòng được chỉ định với giới hạn và
  cách xuống dòng đã xác nhận; hoặc giữ trống/cảnh báo.
- **Ảnh hưởng:** Có thể hoãn nếu người dùng xác nhận AX là tùy chọn và chấp
  thuận không in/giữ cảnh báo khi có dữ liệu. Không tự loại bỏ ghi chú.

## Cách phản hồi

Vui lòng trả lời lần lượt B1–B15. Với mỗi nhóm, có thể:

1. Chọn một phương án nêu trên và cung cấp chi tiết còn thiếu.
2. Gửi quy định, phiếu đã điền chuẩn hoặc bảng giá trị để làm căn cứ.
3. Chỉ định rõ nhóm được để trống/cảnh báo và có thể hoãn.

Việc trả lời tài liệu này chưa tự thay đổi mapping. Sau khi có phê duyệt,
các quy tắc mới chỉ được đưa vào `config/template_mapping.json` sau khi
đối chiếu được bằng chứng và bổ sung kiểm thử. Xuất hộ thật vẫn bị dừng cho
đến khi phiên bản nguồn, xã, năm học và các quyết định bắt buộc được duyệt.
