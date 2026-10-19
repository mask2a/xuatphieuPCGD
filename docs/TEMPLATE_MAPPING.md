# Ánh xạ nguồn sang mẫu phiếu

## Mẫu đã phân tích

- Tệp mẫu duy nhất: `maubaocao.xls`.
- Sheet duy nhất: `Phiếu điều tra 1 mặt`.
- UsedRange: `A1:AG94`; nội dung có nghĩa kết thúc ở hàng 87. Ô `AG94` chỉ
  chứa khoảng trắng.
- Khổ giấy A4 ngang, `Zoom=68`; `PrintArea` chưa được đặt trong tệp gốc.
- Có 442 vùng ô gộp. Không xóa hàng trên workbook này.

## Đầu phiếu và các khối thành viên

| Trường trên phiếu | Ô neo / vùng gộp | Ứng viên nguồn |
|---|---|---|
| Xã (1) | `A1:G1` | Không có cột xã trong `MauNhapLieu`; mẫu đang ghi sẵn `VĂN HÁN` |
| Số phiếu (4) | `Z1:AG1` | `N` |
| Diện cư trú (5) | `Z2:AG2` | `O`; phiếu chỉ hiện `Thường trú` và `Tạm trú` |
| Xóm (2) | `A3:G3` | `M` |
| Họ tên chủ hộ (3) | `H3:Y3` | `K` + `L` |
| Điện thoại (6) | `Z3:AG3` | Chỉ số `AW` của thành viên duy nhất có `AU` xác định là chủ hộ |

Vị trí thành viên đã đối chiếu với các số thứ tự in sẵn:

| Vị trí | Hàng bắt đầu | Năm hàng của khối |
|---:|---:|---|
| 1 | 9 | 9–13 |
| 2 | 14 | 14–18 |
| 3 | 19 | 19–23 |
| 4 | 24 | 24–28 |
| 5 | 29 | 29–33 |
| 6 | 34 | 34–38 |
| 7 | 39 | 39–43 |
| 8 | 47 | 47–51 |
| 9 | 52 | 52–56 |
| 10 | 57 | 57–61 |
| 11 | 62 | 62–66 |
| 12 | 67 | 67–71 |
| 13 | 72 | 72–76 |
| 14 | 77 | 77–81 |

Hàng 44–46 là khoảng trống giữa vị trí 7 và 8. Phần xác nhận bắt đầu hàng 82:
tên/chữ ký ở cột A, các khối ngày ở `F82:I87`, `J82:M87`, `N82:Q87`,
`R82:V87`, `W82:Z87`; nội dung xác nhận UBND ở `AA82:AG83` (vùng gộp).

Với hàng bắt đầu `s` của một thành viên, cấu trúc đã xác minh:

| Trường | Ô đích |
|---|---|
| Số TT mẫu | `A{s}` (số in sẵn, không ghi đè) |
| Họ tên | `B{s}:E{s}` (ô gộp) |
| Quan hệ với chủ hộ | `B{s+1}:E{s+1}` (ô gộp, đang có nhãn) |
| Ngày sinh | `D{s+2}:E{s+2}` (ô gộp, nhãn `Ngày sinh:` ở cột B:C) |
| Nữ | `C{s+3}` (có ký tự mẫu `q`) |
| Dân tộc | `D{s+3}` (nhãn `DT:____`) |
| Tôn giáo | `E{s+3}` (nhãn `TG:_____`) |
| Cha/mẹ/người đỡ đầu | `B{s+4}:E{s+4}` (ô gộp, đang có nhãn) |
| Năm học in sẵn | `F{s}` đến `F{s+4}`; 2025–2026 đến 2029–2030 |
| Lớp/khối từng năm | `H{s+i}:J{s+i}` (ô gộp; `i=0..4`) |
| Trường-huyện từng năm | `K{s+i}:O{s+i}` (ô gộp; `i=0..4`) |
| Tốt nghiệp phổ thông | Cột P (cấp học), Q (bổ túc), R (năm học), trong khối `s..s+4` |
| Tốt nghiệp nghề | Cột S (bậc), T:U (năm học/vùng gộp theo mẫu), trong khối `s..s+4` |
| Học xong / năm | V / W, trong khối `s..s+4` |
| Bỏ học / năm | X / Y, trong khối `s..s+4` |
| Đang học / hoàn thành / tái mù | Z / AA / AB, trong khối `s..s+4` |
| Loại khuyết tật / chứng nhận / khả năng | AC / AD / AE, trong khối `s..s+4` |
| Chuyển đến, chuyển đi, chết | AF, trong khối `s..s+4` |
| Ghi chú | AG, trong khối `s..s+4` |

## Bảng ánh xạ và mức độ chắc chắn

| Nguồn | Ô đích | Đánh giá |
|---|---|---|
| `B`, `C` | `B{s}:E{s}` | Có thể kết hợp họ đệm và tên bằng một dấu cách; không sửa chuỗi thành phần gốc |
| `D`, `E`, `F` | `D{s+2}:E{s+2}` | Các thành phần ngày sinh rõ ràng; giữ ngày/tháng 2 chữ số theo định dạng nguồn; không bổ sung thành phần trống |
| `G` | `C{s+3}` | Nguồn `X` nghĩa là nữ; ký tự `q` và font của ô mẫu cần được kiểm tra khi chọn cách đánh dấu |
| `H` | `D{s+3}` | Trường dân tộc rõ ràng; giữ nhãn DT trên mẫu |
| `I` | `E{s+3}` | Trường tôn giáo rõ ràng; giữ nhãn TG; nguồn mẫu hiện rỗng |
| `K`, `L` | `H3:Y3` | Họ tên chủ hộ; cần cảnh báo mã có nhiều tên chủ hộ |
| `M` | `A3:G3` | Ứng viên cho Xóm/địa chỉ; ý nghĩa `M` rộng hơn nhãn "Xóm" trên phiếu |
| `N` | `Z1:AG1` | Mã số phiếu; giữ nguyên chuỗi |
| `O` | `Z2:AG2` | Ứng viên cho diện cư trú; cần quy tắc đánh dấu checkbox và xử lý khác chữ hoa/thường |
| `P` | `AF` trong khối thành viên | Nguồn là tình trạng cư trú, nhãn đích là chuyển đến/chuyển đi/chết; chỉ ánh xạ sau khi đối chiếu giá trị cụ thể |
| `Q`, `R` | `H:J` một dòng trong khối | Chưa xác định dòng năm học; nguồn không có năm học đi kèm |
| `S`, `T` | `K:O` một dòng trong khối | Hai trường nguồn cùng ứng viên với một ô đích; cần quy tắc ghép trường và huyện |
| `U` | Không có | Mẫu không có ô mã trường |
| `V`, `W`, `X` | `P`, `Q`, `R` trong khối | Nhãn gần nhau; chưa xác định dòng đích cho dữ liệu tốt nghiệp |
| `Y`, `AA` | `S`, `T:U` trong khối | Ứng viên tốt nghiệp nghề; chưa xác định dòng đích |
| `Z` | Không có | Số bằng tốt nghiệp không có ô riêng |
| `AB`, `AC` | `V`, `W` trong khối | Nhãn gần nhau; chưa xác định dòng đích |
| `AD`, `AE` | `X`, `Y` trong khối | Nhãn gần nhau; chưa xác định dòng đích |
| `AF`, `AG`, `AH` | `Z`, `AA`, `AB` trong khối | Nhãn gần nhau; chưa xác định dòng đích |
| `AI:AP` | `AC` trong khối | Nhiều cột loại khuyết tật và một ô "Loại"; chưa có quy tắc biểu diễn nhiều giá trị |
| `AQ`, `AR` | `AD`, `AE` trong khối | Tên trường gần khớp với nhãn phiếu |
| `AS`, `AT` | Không có | Phiếu không có trường hoàn cảnh đặc biệt |
| `AU` | `B{s+1}:E{s+1}` | Quan hệ với chủ hộ |
| `AV` | `B{s+4}:E{s+4}` | Cha/mẹ/người đỡ đầu |
| `AW` | `Z3:AG3` | Chỉ dùng điện thoại của chủ hộ duy nhất theo `AU`; giữ dạng hiển thị/số 0 đầu |
| `AX` | `AG` trong khối thành viên | Ghi chú; cần chốt dòng đặt giá trị trong khối 5 dòng |

## Ánh xạ được phép dùng ở Giai đoạn 3

Ánh xạ máy chạy nằm trong `../config/template_mapping.json`. Chỉ các ô đã
xác minh và có quy tắc bảo toàn nhãn mới được ghi:

| Nguồn/giá trị | Ô đích | Quy tắc |
|---|---|---|
| `commune_name` từ cấu hình | `A1:G1` và tên xã ở `AA82:AG83` | Thay xã mẫu bằng cấu hình; thay riêng đoạn tên xã trong nội dung xác nhận để giữ định dạng chữ hỗn hợp |
| Mã hộ `N` | `Z1:AG1` | Thay phần gạch trống sau nhãn; giữ chuỗi và số 0 đầu |
| Tên chủ hộ `K:L` | `H3:Y3` | Chỉ ghi khi nhất quán trong hộ; mâu thuẫn thì giữ ô mẫu và cảnh báo |
| Điện thoại chủ hộ `AW` | `Z3:AG3` | Chỉ ghi số đã xác định duy nhất qua `AU`; không lấy số thành viên khác |
| Họ tên thành viên `B:C` | `B{s}:E{s}` | Họ tên đã chuẩn hóa |
| Quan hệ `AU` | `B{s+1}:E{s+1}` | Giữ nhãn mẫu `QH với chủ hộ:` |
| Ngày/tháng/năm `D:F` | `D{s+2}:E{s+2}` | Giữ vị trí thành phần trống và số 0 hiển thị; không tự điền thành phần thiếu |
| Dân tộc `H` | `D{s+3}` | Giữ nhãn `DT:` |
| Tôn giáo `I` | `E{s+3}` | Giữ nhãn `TG:` |
| Cha/mẹ/người đỡ đầu `AV` | `B{s+4}:E{s+4}` | Giữ nhãn mẫu |

Chỉ phần dữ liệu sau nhãn được thay; các ô đích được kiểm tra không có công
thức trước khi ghi. Tệp mẫu được mở chỉ đọc, sao chép bằng Excel COM và không
bị sửa. Workbook lưu thành `.xlsx`, mở lại bằng Excel để kiểm tra số thành
viên, vùng gộp, định dạng ô đại diện, chiều rộng cột, chiều cao hàng, thiết
lập trang, vùng in, phần xác nhận và external links.

Các trường còn lại được khai báo trong mapping JSON là **chưa ánh xạ**. Nếu
có dữ liệu nguồn, báo cáo xuất ghi cảnh báo và giữ nguyên mẫu; không ghi
đoán vào ô đích:

- Địa chỉ `M`, diện cư trú `O`, giới tính `G`: đích/biểu tượng đánh dấu chưa
  được xác minh đầy đủ.
- Diện ưu tiên `J`: chưa có trường riêng tương ứng đã xác minh trên mẫu;
  nếu nguồn có giá trị, exporter ghi cảnh báo thay vì bỏ qua.
- Tình trạng cư trú `P`, học tập/trường `Q:T`, tốt nghiệp và quá trình học
  `V:AH`: chưa có quy tắc gắn dòng theo năm học.
- Khuyết tật `AI:AP`, chứng nhận/khả năng học `AQ:AR`, hoàn cảnh `AS:AT`,
  ghi chú `AX`: chưa chốt cách biểu diễn hoặc dòng đích.
- Các dòng năm học có sẵn trong mẫu giữ nguyên; `school_year` hiện chưa được
  dùng để sửa nội dung cố định của biểu mẫu.

Do đó, Giai đoạn 3 hiện chỉ xuất các trường đã xác minh. Chưa thể coi toàn
bộ ánh xạ dữ liệu là hoàn tất cho nghiệp vụ thực tế cho tới khi các quy tắc
chưa ánh xạ được người dùng xác nhận.

Trước khi xuất, exporter đọc `config/settings.json` (hoặc đường dẫn cấu hình
được truyền rõ), yêu cầu `commune_name` và `school_year` hợp lệ, và kiểm tra
địa bàn từng hộ khớp cấu hình. Năm học có dạng `YYYY-YYYY` với hai năm liên
tiếp. Giá trị năm học được nêu trong cảnh báo; chưa thay đổi các dòng năm in
sẵn khi chưa có quy tắc gắn trường nguồn vào đúng năm.

Không được ghi đè các nhãn có sẵn (`QH với chủ hộ:`, `Ngày sinh:`,
`Cha, mẹ, Ng đỡ đầu:`, `DT:`, `TG:`) mà không bảo toàn nhãn. Không thể coi
những ô trống là lý do để tự điền giá trị mặc định.

### Quy tắc điện thoại đã chốt

- Xác định chủ hộ từ quan hệ ở cột `AU`; chỉ khi xác định được đúng một bản
  ghi chủ hộ mới lấy điện thoại tương ứng ở `AW`.
- Không dùng điện thoại của thành viên khác làm giá trị thay thế. Nếu bản
  ghi chủ hộ duy nhất không có điện thoại, để trống.
- Nếu không có hoặc có nhiều bản ghi chủ hộ, để trống và cảnh báo. Nếu nhiều
  bản ghi chủ hộ có số khác nhau, ghi thêm cảnh báo mâu thuẫn; không tự chọn.
- Giữ chuỗi số điện thoại và số 0 đầu theo giá trị hiển thị của Excel. Việc
  các thành viên không phải chủ hộ có những số điện thoại khác nhau không
  phải là lỗi hộ.

## Thử nghiệm đề xuất phiếu ngắn

Đã tạo bản sao tạm thời của `maubaocao.xls` và thử trên bản sao cho hộ 1 và
6 thành viên; tệp mẫu gốc không bị sửa. Các thao tác trên bản sao:

1. Ẩn các hàng từ vị trí thành viên tiếp theo đến 81: hộ 1 ẩn `14:81`, hộ 6
   ẩn `39:81`. Không xóa hàng và không sửa vùng gộp.
2. Đặt `PrintArea` là `$A$1:$AG$87`, để gồm phần xác nhận và loại hàng 88–94.
3. Đặt `Zoom=False`, `FitToPagesWide=1`, `FitToPagesTall=1`. Trong tệp gốc,
   `Zoom=68` đang được đặt đồng thời với FitToPages, nên cần tắt Zoom để các
   tham số fit có tác dụng.
4. Xuất PDF từ bản sao: cả hai trường hợp cho 1 trang. Kiểm tra nội dung xác
   nhận tại `AA82:AG83` còn nguyên, vùng gộp này không bị ẩn; đếm vùng gộp
   vẫn là 442.

Kết luận: ẩn hàng không sử dụng an toàn hơn xóa hàng và đã đạt được bố cục
ngắn một trang trên các bản thử. Trước khi đưa vào exporter vẫn cần kiểm tra
hiển thị thực tế trên Excel/bản in, đặc biệt kích thước chữ sau khi fit một
trang và cách điền checkbox.
