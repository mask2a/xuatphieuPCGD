# Audit cảnh báo và ánh xạ biểu mẫu — Giai đoạn 3

## Phạm vi và kết luận

Audit đối chiếu `docs/DATA_ANALYSIS.md`, `docs/TEMPLATE_MAPPING.md`,
`docs/SPEC.md`, `config/template_mapping.json` và bộ xuất Excel với
`maubaocao.xls`. Mẫu và dữ liệu nguồn được mở chỉ đọc.

**15 cảnh báo trong báo cáo kiểm thử trước là 15 lần phát cảnh báo, không
phải 15 loại trường khác nhau.** Chúng là ba cảnh báo lặp lại cho từng hộ
kiểm thử tổng hợp `TEST_01`, `TEST_06`, `TEST_07`, `TEST_14`, `TEST_15`.
Không được hiểu kết quả 5/5 workbook là bằng chứng các trường nghiệp vụ đã
được ánh xạ đầy đủ.

Không bổ sung đích ghi mới trong audit này: các vị trí còn lại mới chỉ là
ứng viên hoặc chưa xác định được cách biểu diễn. Thêm `priority_category`
(cột J) vào danh sách chưa ánh xạ để không bỏ qua dữ liệu nếu trường này
xuất hiện ở nguồn.

## Chi tiết 15 cảnh báo trong báo cáo

| # | Hộ kiểm thử | Mã cảnh báo | Trường nguồn / ý nghĩa | Ô đích trên mẫu | Phân loại và lý do |
|---:|---|---|---|---|---|
| 1 | `TEST_01` | `HOUSEHOLD_RESIDENCE_NOT_EXPORTED` | O — diện cư trú cấp hộ | Ứng viên `Z2:AG2` | Có trường trên mẫu; kiểu đánh dấu/giá trị chưa xác minh |
| 2 | `TEST_01` | `SCHOOL_YEAR_ROWS_NOT_MAPPED` | Các dữ liệu học tập, tốt nghiệp, quá trình học gắn với năm | Các dải năm in sẵn; chưa chọn hàng đích | Quy tắc gắn năm học chưa xác minh |
| 3 | `TEST_01` | `UNMAPPED_FEMALE` | G — X nghĩa là nữ | Ứng viên `C{s+3}` | Có vị trí; chưa xác minh ký tự Wingdings để đánh dấu |
| 4 | `TEST_06` | `HOUSEHOLD_RESIDENCE_NOT_EXPORTED` | O — diện cư trú cấp hộ | Ứng viên `Z2:AG2` | Có trường trên mẫu; kiểu đánh dấu/giá trị chưa xác minh |
| 5 | `TEST_06` | `SCHOOL_YEAR_ROWS_NOT_MAPPED` | Các dữ liệu học tập, tốt nghiệp, quá trình học gắn với năm | Các dải năm in sẵn; chưa chọn hàng đích | Quy tắc gắn năm học chưa xác minh |
| 6 | `TEST_06` | `UNMAPPED_FEMALE` | G — X nghĩa là nữ | Ứng viên `C{s+3}` | Có vị trí; chưa xác minh ký tự Wingdings để đánh dấu |
| 7 | `TEST_07` | `HOUSEHOLD_RESIDENCE_NOT_EXPORTED` | O — diện cư trú cấp hộ | Ứng viên `Z2:AG2` | Có trường trên mẫu; kiểu đánh dấu/giá trị chưa xác minh |
| 8 | `TEST_07` | `SCHOOL_YEAR_ROWS_NOT_MAPPED` | Các dữ liệu học tập, tốt nghiệp, quá trình học gắn với năm | Các dải năm in sẵn; chưa chọn hàng đích | Quy tắc gắn năm học chưa xác minh |
| 9 | `TEST_07` | `UNMAPPED_FEMALE` | G — X nghĩa là nữ | Ứng viên `C{s+3}` | Có vị trí; chưa xác minh ký tự Wingdings để đánh dấu |
| 10 | `TEST_14` | `HOUSEHOLD_RESIDENCE_NOT_EXPORTED` | O — diện cư trú cấp hộ | Ứng viên `Z2:AG2` | Có trường trên mẫu; kiểu đánh dấu/giá trị chưa xác minh |
| 11 | `TEST_14` | `SCHOOL_YEAR_ROWS_NOT_MAPPED` | Các dữ liệu học tập, tốt nghiệp, quá trình học gắn với năm | Các dải năm in sẵn; chưa chọn hàng đích | Quy tắc gắn năm học chưa xác minh |
| 12 | `TEST_14` | `UNMAPPED_FEMALE` | G — X nghĩa là nữ | Ứng viên `C{s+3}` | Có vị trí; chưa xác minh ký tự Wingdings để đánh dấu |
| 13 | `TEST_15` | `HOUSEHOLD_RESIDENCE_NOT_EXPORTED` | O — diện cư trú cấp hộ | Ứng viên `Z2:AG2` | Có trường trên mẫu; kiểu đánh dấu/giá trị chưa xác minh |
| 14 | `TEST_15` | `SCHOOL_YEAR_ROWS_NOT_MAPPED` | Các dữ liệu học tập, tốt nghiệp, quá trình học gắn với năm | Các dải năm in sẵn; chưa chọn hàng đích | Quy tắc gắn năm học chưa xác minh |
| 15 | `TEST_15` | `UNMAPPED_FEMALE` | G — X nghĩa là nữ | Ứng viên `C{s+3}` | Có vị trí; chưa xác minh ký tự Wingdings để đánh dấu |

## Phân loại 36 trường chưa ánh xạ A/B/C

Tiêu chí:

- **A — Đủ căn cứ để ánh xạ:** ô đích cụ thể và quy tắc nghiệp vụ/biểu
  diễn đều đã được xác minh trực tiếp. Có tọa độ ứng viên nhưng thiếu quy
  tắc không đủ điều kiện A.
- **B — Có khả năng ánh xạ, cần xác nhận:** trên phiếu có vùng/nhãn gần
  tương ứng, nhưng thiếu quy tắc nghiệp vụ, hàng năm đích, cách biểu diễn,
  hoặc giải quyết xung đột.
- **C — Không có vị trí phù hợp:** mẫu không có ô riêng cho trường đó;
  không được nhét dữ liệu vào một ô khác.

| Nhóm | Số trường | Trường |
|---|---:|---|
| A | 0 | Không có trường chưa ánh xạ nào vừa có ô đích vừa có quy tắc ghi đã xác minh |
| B | 31 | `address`, `residence`, `female`, `residence_status`, `grade_group`, `class_name`, `school_district`, `school_name`, `graduation_level`, `supplementary_education`, `graduation_year`, `vocational_graduation_level`, `vocational_graduation_year`, `completed_education`, `completed_education_year`, `dropped_out_class`, `dropped_out_year`, `currently_studying_class`, `completed_class`, `reliteracy_level`, `disability_mobility`, `disability_hearing_speech`, `disability_vision`, `disability_neurological`, `disability_intellectual`, `disability_learning`, `disability_autism`, `disability_other`, `disability_certified`, `learning_capacity`, `notes` |
| C | 5 | `priority_category`, `school_code`, `diploma_number`, `special_circumstances`, `special_circumstances_details` |

Danh mục chi tiết dưới đây nêu tọa độ ứng viên và lý do. Với nhóm B, tọa độ
quan sát được không phải quyền ghi dữ liệu. Với nhóm C, không có đích phù
hợp đã xác minh trên mẫu hiện tại.

## Bảng quyết định cho 31 trường nhóm B

Mức ưu tiên dưới đây là **đề xuất phục vụ hỏi/duyệt**, không phải kết luận
về tính bắt buộc theo quy định. Requiredness nghiệp vụ hiện chưa được xác
nhận. Các trường được gom theo câu hỏi chung để giảm số lượt xác nhận.

| Nhóm/câu hỏi | Trường nguồn | Cột | Ý nghĩa nghiệp vụ | Vị trí đích có khả năng tương ứng | Quy tắc còn thiếu | Ưu tiên đề xuất |
|---|---|---|---|---|---|---|
| B1 | `female` | G | X biểu thị nữ | `C{s+3}`, cạnh nhãn “Nữ:” | Ký tự/định dạng chính xác để đánh dấu; cách xử lý nguồn khác `X` hoặc trống | Cần thiết; requiredness chưa xác nhận |
| B2 | `residence` | O | Diện cư trú cấp hộ | `Z2:AG2`, có lựa chọn Thường trú/Tạm trú | Quy tắc chọn đúng placeholder; chuẩn hóa viết hoa/thường; xử lý xung đột trong hộ | Cần thiết; requiredness chưa xác nhận |
| B3 | `address` | M | Địa chỉ gồm số nhà/đường/tổ nếu có | `A3:G3`, nhãn “Xóm (2)” | Có được in toàn bộ địa chỉ vào ô “Xóm” hay chỉ một thành phần; xử lý độ dài | Cần thiết; requiredness chưa xác nhận |
| B4 | `residence_status` | P | Tình trạng cư trú cá nhân | Ứng viên `AF` trong khối, nhãn trạng thái chuyển đến/đi/chết | Từ điển giá trị P và phép tương đương với nhãn đích; không đồng nhất với O | Cần thiết nếu có dữ liệu; requiredness chưa xác nhận |
| B5 | `grade_group`, `class_name` | Q, R | Khối/lớp học hiện tại | Ứng viên `H:J` tại một hàng trong khối thành viên | Hàng năm đích khi nguồn không kèm năm; quy tắc ghép khối/lớp trong vùng đích | Cần thiết nếu có dữ liệu; requiredness chưa xác nhận |
| B6 | `school_district`, `school_name` | S, T | Huyện và tên trường đang học | Ứng viên `K:O` ở hàng tương ứng của khối | Thứ tự/chuỗi kết hợp hai trường; delimiter; cách xử lý một trường bị thiếu; hàng năm | Cần thiết nếu có dữ liệu; requiredness chưa xác nhận |
| B7 | `graduation_level`, `supplementary_education`, `graduation_year` | V, W, X | Bậc tốt nghiệp phổ thông, bổ túc và năm tốt nghiệp | Ứng viên P, Q, R trong khối | Hàng năm đích; giá trị nào được ghi cho từng loại tốt nghiệp; định dạng năm | Cần thiết khi nguồn có dữ liệu; requiredness chưa xác nhận |
| B8 | `vocational_graduation_level`, `vocational_graduation_year` | Y, AA | Bậc và năm tốt nghiệp nghề | Ứng viên S và T:U trong khối | Hàng năm; định dạng năm; vùng T:U và cách ghi khi thiếu một thành phần | Cần thiết khi nguồn có dữ liệu; requiredness chưa xác nhận |
| B9 | `completed_education`, `completed_education_year` | AB, AC | Học xong và năm hoàn thành | Ứng viên V, W trong khối | Hàng năm, quan hệ giữa trạng thái học xong và năm ghi nhận | Cần thiết khi nguồn có dữ liệu; requiredness chưa xác nhận |
| B10 | `dropped_out_class`, `dropped_out_year` | AD, AE | Bỏ học và năm bỏ học | Ứng viên X, Y trong khối | Hàng năm; xác định giá trị/lớp và định dạng năm; xử lý dữ liệu thiếu cặp | Cần thiết khi nguồn có dữ liệu; requiredness chưa xác nhận |
| B11 | `currently_studying_class`, `completed_class` | AF, AG | Lớp đang học và lớp hoàn thành | Ứng viên Z, AA trong khối | Hàng năm; phân biệt/ưu tiên hai trạng thái nếu cùng có dữ liệu | Cần thiết khi nguồn có dữ liệu; requiredness chưa xác nhận |
| B12 | `reliteracy_level` | AH | Mức tái mù chữ | Ứng viên AB trong khối | Hàng năm; từ điển giá trị và cách ghi mức | Cần thiết nếu có trường hợp; requiredness chưa xác nhận |
| B13 | `disability_mobility`, `disability_hearing_speech`, `disability_vision`, `disability_neurological`, `disability_intellectual`, `disability_learning`, `disability_autism`, `disability_other` | AI:AP | Các loại khuyết tật | Ứng viên AC “Loại” trong khối | Cách ghép nhiều loại vào một ô; thứ tự/delimiter; cách biểu diễn “khác”; hàng năm | Cần thiết nếu có trường hợp; requiredness chưa xác nhận |
| B14 | `disability_certified`, `learning_capacity` | AQ, AR | Chứng nhận khuyết tật và khả năng học tập | Ứng viên AD, AE trong khối | Giá trị được phép, hàng/năm đích, tương quan với loại khuyết tật | Cần thiết nếu có trường hợp; requiredness chưa xác nhận |
| B15 | `notes` | AX | Ghi chú thành viên | Ứng viên AG trong khối | Chọn hàng nào trong khối 5 dòng; quy tắc độ dài/xuống dòng | Tùy chọn đề xuất; chưa được phê duyệt |

### Câu hỏi xác nhận gom theo quy tắc

1. **B1 — Giới tính:** Khi cột G có `X`, chính xác ô/ký tự nào phải được đánh
   dấu tại `C{s+3}`? Có thể xác nhận bằng ảnh/phiếu đã điền chuẩn hoặc hướng
   dẫn biểu mẫu nào? Giá trị G trống/khác X phải để trống hay có quy tắc khác?
2. **B2 — Diện cư trú:** Với O bằng Thường trú/Tạm trú (kể cả khác hoa
   thường), placeholder nào ở `Z2` được đánh dấu và bằng ký tự gì? Nếu hộ có
   nhiều giá trị O mâu thuẫn, xác nhận tiếp tục để trống/cảnh báo hay có quy
   trình duyệt thủ công?
3. **B3 — Địa chỉ:** Cột M có được in nguyên văn vào ô “Xóm (2)” không, hay
   chỉ lấy thành phần tổ/xóm? Nếu nguồn chứa số nhà/đường mà không có xóm,
   cần để trống hay vẫn in nguyên địa chỉ? Có giới hạn/chuẩn rút gọn nào?
4. **B4 — Tình trạng cư trú:** Xin cung cấp danh sách đầy đủ giá trị có thể
   có ở P và ý nghĩa từng giá trị; giá trị nào tương ứng các nhãn đích ở AF
   (chuyển đến/chuyển đi/chết)? Các giá trị khác phải để trống/cảnh báo?
5. **B5 — Khối/lớp hiện tại:** Q/R không có năm học đi kèm; hàng nào trong
   5 dòng năm được dùng cho dữ liệu hiện tại? Có ghi cả khối và lớp vào
   `H:J` không, theo định dạng nào? Xác nhận quan hệ với cột AF nếu AF cũng
   chứa lớp đang học.
6. **B6 — Huyện/trường:** Có được ghép S và T vào `K:O` không? Xin chốt thứ
   tự, dấu phân cách, cách xử lý thiếu huyện/trường và hàng năm dùng cho dữ
   liệu trường hiện tại.
7. **B7–B12 — Học tập/tốt nghiệp:** Xin cung cấp bảng đối chiếu mỗi nguồn
   với nhãn đích và hàng/năm. Với X/AA/AC/AE có năm tường minh, năm đó được
   ánh xạ vào dòng nào trong các năm F9:F13? Với Q/R và các trạng thái không
   kèm năm, dùng năm học cấu hình hay để trống/cảnh báo? Nếu nhiều kết quả
   cùng tồn tại, ghi tất cả hay chỉ một kết quả theo quy tắc nào?
8. **B13–B14 — Khuyết tật:** Xin xác nhận giá trị nguồn nào được coi là có
   khuyết tật/chứng nhận/khả năng học; cách ghi khi có nhiều loại trong ô
   AC; delimiter, thứ tự và hàng/năm cho AQ/AR.
9. **B15 — Ghi chú:** Ghi AX vào hàng nào trong khối 5 hàng? Có được xuống
   dòng/giới hạn độ dài hay không? Nếu không có quy tắc, xác nhận tiếp tục
   cảnh báo và không ghi.

Các mức ưu tiên trên không tự biến trường thành required. Cần người phụ
trách nghiệp vụ xác nhận trường nào là bắt buộc, cần thiết hoặc tùy chọn
trước nghiệm thu.

## Danh mục trường chưa ánh xạ

`đích ứng viên` bên dưới chỉ ghi nhận cấu trúc in quan sát được, không phải
quyền ghi dữ liệu. Các trường học/tốt nghiệp nằm trong khối 5 dòng của một
thành viên; chọn dòng phải dựa trên quy tắc năm học đã được xác nhận.

| Nguồn | Ý nghĩa theo tiêu đề nguồn | Ô đích đã quan sát / ứng viên | Trạng thái nghiệp vụ |
|---|---|---|---|
| J `priority_category` | Diện ưu tiên | Không có trường riêng đã xác minh | Không có đích; optional/required chưa được quy định |
| M `address` | Địa chỉ: số nhà, đường, tổ | `A3:G3`, nhãn “Xóm (2)” | Chưa xác nhận có được rút gọn/hiển thị địa chỉ vào ô Xóm |
| O `residence` | Diện cư trú cấp hộ | `Z2:AG2` | Có trường; chưa xác minh lựa chọn checkbox, ký hiệu hoặc xử lý biến thể |
| G `female` | `X` biểu thị nữ | `C{s+3}` | Vị trí ứng viên; chưa xác minh ký tự Wingdings/biểu diễn giới tính |
| P `residence_status` | Tình trạng cư trú | `AF` trong khối thành viên | Giá trị nguồn (ví dụ “Lưu Trú”) chưa đối chiếu với nhãn chuyển đến/đi/chết |
| Q `grade_group` | Khối học | `H:J` trên một hàng năm | Chưa có hàng năm đích |
| R `class_name` | Lớp học | `H:J` trên một hàng năm | Chưa xác nhận tương quan với khối và năm đích |
| S `school_district` | Quận/huyện đang học | Có thể cùng `K:O` với tên trường | Chưa có quy tắc ghép; chỉ một giá trị khác rỗng trong nguồn mẫu |
| T `school_name` | Tên trường đang học | `K:O` trên một hàng năm | Chưa có quy tắc ghép với huyện hoặc hàng năm |
| U `school_code` | Mã trường | Không có ô mã trường riêng | Không có đích trên mẫu |
| V `graduation_level` | Bậc tốt nghiệp | Ứng viên cột P trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| W `supplementary_education` | Bổ túc | Ứng viên cột Q trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| X `graduation_year` | Năm tốt nghiệp | Ứng viên cột R trong khối | Chưa xác nhận hàng năm và cách biểu diễn |
| Y `vocational_graduation_level` | Bậc tốt nghiệp nghề | Ứng viên cột S trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| AA `vocational_graduation_year` | Năm tốt nghiệp nghề | Ứng viên cột T:U trong khối | Chưa xác nhận hàng năm và cách biểu diễn |
| Z `diploma_number` | Số bằng tốt nghiệp | Không có ô số bằng riêng | Không có đích trên mẫu |
| AB `completed_education` | Học xong | Ứng viên cột V trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| AC `completed_education_year` | Năm hoàn thành | Ứng viên cột W trong khối | Chưa xác nhận hàng năm và cách biểu diễn |
| AD `dropped_out_class` | Bỏ học | Ứng viên cột X trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| AE `dropped_out_year` | Năm bỏ học | Ứng viên cột Y trong khối | Chưa xác nhận hàng năm và cách biểu diễn |
| AF `currently_studying_class` | Đang học lớp | Ứng viên cột Z trong khối | Chưa xác nhận ý nghĩa với Q/R và hàng năm |
| AG `completed_class` | Hoàn thành lớp | Ứng viên cột AA trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| AH `reliteracy_level` | Tái mù chữ mức | Ứng viên cột AB trong khối | Chưa xác nhận hàng năm và quy tắc chọn |
| AI `disability_mobility` | Khuyết tật vận động | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AJ `disability_hearing_speech` | Khuyết tật nghe, nói | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AK `disability_vision` | Khuyết tật nhìn | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AL `disability_neurological` | Khuyết tật thần kinh, tâm thần | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AM `disability_intellectual` | Khuyết tật trí tuệ | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AN `disability_learning` | Khuyết tật học tập | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AO `disability_autism` | Tự kỷ | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AP `disability_other` | Khuyết tật khác | Cột AC “Loại” | Nhiều loại nguồn dùng chung một ô; chưa có quy tắc ghép |
| AQ `disability_certified` | Có chứng nhận khuyết tật | Ứng viên cột AD trong khối | Chưa xác nhận hàng năm/biểu diễn giá trị |
| AR `learning_capacity` | Khả năng học tập | Ứng viên cột AE trong khối | Chưa xác nhận hàng năm/biểu diễn giá trị |
| AS `special_circumstances` | Hoàn cảnh đặc biệt | Không có trường tương ứng đã xác minh | Không có đích trên mẫu |
| AT `special_circumstances_details` | Chi tiết hoàn cảnh | Không có trường tương ứng đã xác minh | Không có đích trên mẫu |
| AX `notes` | Ghi chú thành viên | Ứng viên cột AG trong khối | Chưa xác định hàng của khối 5 dòng |

### Bằng chứng trực tiếp trên mẫu cho G, J, O và các dòng năm

Mẫu được mở bằng Excel COM ở chế độ `ReadOnly=True`; sau khi đóng workbook,
SHA-256 của `maubaocao.xls` trước/sau không đổi.

| Trường | Quan sát trong Excel | Kết luận audit |
|---|---|---|
| G — Nữ | `B12` ghi `Nữ:`; `C12` có giá trị `q`, font Wingdings | Xác nhận có ô/placeholder ứng viên `C{s+3}`. Chưa xác minh glyph/giá trị nào biểu thị lựa chọn hay cách mã hóa khác của nguồn; xếp B, không ghi |
| J — Diện ưu tiên | Không thấy trường riêng dành cho diện ưu tiên trên phiếu; danh mục vùng thành viên không có nhãn tương ứng | Xếp C; giữ cảnh báo nếu dữ liệu nguồn có giá trị |
| O — Diện cư trú | `Z2` có nội dung `Diện cư trú (5): q Thường trú   q Tạm trú` | Xác nhận hai lựa chọn in ở ô đầu phiếu, không xác nhận ký hiệu đánh dấu lựa chọn hay chuẩn hóa giá trị nguồn; xếp B, không ghi |
| Năm học | `F9:F13` in lần lượt `2025 - 2026` đến `2029 - 2030`; hàng lớp/trường và tốt nghiệp có cột trống trong khối thành viên | Xác nhận cấu trúc nhiều hàng/năm trên phiếu. Nguồn không cung cấp khóa năm cho từng thông tin học tập; chưa thể gắn dòng từ `school_year` mà không suy đoán; xếp B |
| Học tập/tốt nghiệp | Các cột ứng viên H:AE nằm trên cùng các hàng với nhãn năm cố định | Tọa độ theo cột có thể nhận biết, nhưng quy tắc chọn hàng/năm và tương quan giữa các trường chưa được xác nhận; xếp B |

Nguồn hiện có giá trị `female` ở 185 bản ghi, `residence` ở 427 bản ghi,
`priority_category` không có giá trị khác rỗng; các nhóm học tập/tốt nghiệp
có nhiều giá trị. Việc J trống trong tệp mẫu không chứng minh J là trường
tùy chọn trong mọi nghiệp vụ.

## Phân loại mức độ bắt buộc

| Nhóm | Trường | Căn cứ và xử lý |
|---|---|---|
| Bắt buộc để cho phép xuất | `commune_name`, `school_year` | Cấu hình ngoài; thiếu hoặc sai định dạng thì dừng trước khi tạo thư mục đầu ra/khởi chạy COM |
| Bắt buộc để nhận diện phiếu/người | Mã hộ N, họ tên thành viên B:C | Cần để xác định phiếu và thành viên; quy tắc xử lý giá trị thiếu trong phiếu chính thức chưa được phê duyệt |
| Được để trống khi nguồn thiếu theo quy tắc đã chốt | Điện thoại AW của chủ hộ duy nhất; từng thành phần ngày sinh D:F | Không dùng dữ liệu thành viên khác để thay điện thoại; không điền thành phần ngày sinh trống |
| Tùy chọn hoặc requiredness chưa được quy định | Dân tộc H, tôn giáo I, quan hệ AU, cha/mẹ/người đỡ đầu AV và các trường đang chờ ánh xạ | Ghi khi có nguồn theo mapping đã được xác minh; không tự tạo giá trị |
| Không có đích trên mẫu | J diện ưu tiên, U mã trường, Z số bằng, AS/AT hoàn cảnh đặc biệt | Không nhét vào trường khác; cần biểu mẫu/đặc tả mới nếu nghiệp vụ bắt buộc |

Các trường có đích ứng viên nhưng chưa rõ requiredness/biểu diễn được đánh
dấu **chưa xác định**, không được phân loại thành optional để hợp thức hóa
việc bỏ qua.

## Dữ liệu thật và cổng cấu hình

Phân tích chỉ đọc `dulieumau.xls` xác nhận 1 tệp đọc, 0 lỗi, 427 bản ghi,
427 thành viên và 120 hộ. Phân bố số thành viên là 18 hộ một người, 15 hộ
hai người, 16 hộ ba người, 36 hộ bốn người, 25 hộ năm người, 9 hộ sáu
người và 1 hộ tám người. Có 3 hộ mâu thuẫn tên chủ hộ, 5 hộ mâu thuẫn diện
cư trú và 10 hộ không xác định được duy nhất chủ hộ. Không chọn tên, diện
cư trú hoặc điện thoại thay thế cho các hộ mâu thuẫn.

`config/settings.json` chưa có giá trị địa bàn/năm học hợp lệ. Vì vậy chưa
xuất workbook từ hộ thật; mẫu và nguồn không bị sửa. `export_households`
đọc cấu hình này (hoặc đường dẫn settings được truyền rõ), yêu cầu
`school_year` theo dạng `YYYY-YYYY` với hai năm liên tiếp, kiểm tra địa bàn
của hộ khớp cấu hình, và dừng trước khi tạo output nếu kiểm tra thất bại.

Năm học được ghi trong cảnh báo export để tránh bị hiểu là đã áp dụng. Các
năm in sẵn trong mẫu và các trường học chưa được thay đổi cho tới khi quy
tắc gắn dữ liệu nguồn vào dòng/năm cụ thể được xác nhận.

Trong lượt audit 2026-10-20, đọc lại nguồn cho kết quả 427 bản ghi/120 hộ;
hash nguồn trước/sau lượt đọc không đổi (`5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`).
Hash này khác hash nguồn ghi trong báo cáo đối chiếu cũ; không suy luận
nguyên nhân hoặc thời điểm khác biệt. Hash `maubaocao.xls` trước/sau kiểm
tra Excel COM chỉ đọc không đổi.
