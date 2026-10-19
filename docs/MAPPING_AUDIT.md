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
