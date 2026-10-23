# PCGD HỘ GIA ĐÌNH

Ứng dụng Windows đọc dữ liệu phổ cập giáo dục từ Excel, nhóm thành viên theo hộ gia đình và xuất phiếu Excel/PDF theo biểu mẫu. Có giao diện tiếng Việt, nhật ký xử lý, xuất nhiều nguồn và gộp PDF theo từng nguồn.

## 1. Chuẩn bị

- Máy Windows có Microsoft Excel bản desktop, đã mở và sử dụng được bình thường.
- Python 3.10 trở lên, có thành phần Tcl/Tk để chạy giao diện Tkinter.
- Tệp dữ liệu nguồn `.xls` hoặc `.xlsx` đúng cấu trúc đang sử dụng.
- Biểu mẫu Excel phù hợp với ánh xạ của dự án, thường đặt tại `templates/maubaocao.xls`.

**Biểu mẫu trống được cung cấp sẵn tại `templates/maubaocao.xls`.** Tệp dữ liệu nguồn và kết quả xuất không được đưa lên GitHub theo `.gitignore`. Khi tải dự án sang máy mới, chỉ cần chuẩn bị dữ liệu nguồn và chọn tệp trong ứng dụng; cấu hình mặc định dùng biểu mẫu đi kèm và thư mục kết quả `output` trong dự án.

## 2. Cài đặt và khởi động

Cách nhanh trên Windows:

1. Cài Python 3.10 trở lên (có Tcl/Tk) và Microsoft Excel bản desktop.
2. Tải ZIP từ GitHub rồi giải nén, hoặc dùng `git clone https://github.com/mask2a/xuatphieuPCGD.git`.
3. Mở `setup.bat` để tạo môi trường ảo, tải thư viện và kiểm tra biểu mẫu/cấu hình. Lần cài đầu cần Internet.
4. Mở `run.bat` để chạy. Nếu chưa có môi trường ảo, file này tự gọi bước cài đặt.
5. Chọn dữ liệu nguồn của bạn trong giao diện, kiểm tra xã/phường và năm học rồi xuất.

Repo đã kèm biểu mẫu, ánh xạ và cấu hình đường dẫn tương đối. Không cần chép thư mục `venv` từ máy khác. Sau khi cập nhật mã bằng `git pull`, chạy lại `setup.bat` để cài thư viện nếu có thay đổi. Chọn thư mục giải nén có quyền ghi để lưu môi trường và kết quả.

Nếu muốn thao tác bằng lệnh:

Mở PowerShell tại thư mục dự án, chạy lần lượt:

```powershell
py -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
.\venv\Scripts\python.exe main.py
```

Nếu đã có môi trường `venv` và đã cài thư viện, những lần sau chỉ cần:

```powershell
.\venv\Scripts\python.exe main.py
```

Các lệnh trên dùng trực tiếp Python trong môi trường ảo, không cần kích hoạt môi trường bằng `Activate.ps1`.

Kiểm tra cài đặt độc lập:

```powershell
.\venv\Scripts\python.exe scripts/check_install.py
```

Công cụ kiểm tra thư viện, Tkinter, cấu hình, sheet biểu mẫu và đăng ký Excel COM; không mở Excel hoặc xuất dữ liệu. Muốn xác nhận xuất thực tế, chạy `run.bat` trong phiên Windows đang đăng nhập rồi xuất một hộ thử.

## 3. Đọc và kiểm tra dữ liệu

1. Bấm **Chọn tệp** để chọn dữ liệu nguồn, hoặc **Chọn thư mục** để đọc nhiều tệp trong thư mục.
2. Chọn **Biểu mẫu Excel** và **Thư mục kết quả**.
3. Nhập đúng **Xã/phường** và **Năm học**, sau đó bấm **Lưu cấu hình** để lưu các lựa chọn.
4. Bấm **Đọc dữ liệu**, chờ danh sách hộ hiện ra.
5. Tìm theo mã hộ hoặc tên chủ hộ. Chọn hộ rồi bấm **Xem thành viên** để rà soát thông tin và cảnh báo.

Cấu hình đọc hiện tại ưu tiên sheet `MauNhapLieu`, dòng tiêu đề 2 và dữ liệu bắt đầu từ dòng 5. Các cột chính gồm B (họ đệm), C (tên), K/L (họ tên chủ hộ), M (địa chỉ) và N (số phiếu/mã hộ). Ánh xạ đầy đủ nằm trong `config/settings.json`.

Với tệp có cấu trúc khác, cần điều chỉnh cấu hình đọc và ánh xạ trước khi xuất. Nút **Cấu hình ánh xạ** dùng để xem thông tin ánh xạ hiện có.

## 4. Xuất phiếu Excel và PDF

### Xuất từ dữ liệu nguồn

1. Đọc và rà soát dữ liệu theo mục 3.
2. Bấm **Xuất hộ đã chọn** để xuất Excel cho một hộ, hoặc **Xuất tất cả Excel** để xuất toàn bộ danh sách đã đọc. Ô tìm kiếm chỉ lọc danh sách hiển thị; nút xuất tất cả vẫn xử lý toàn bộ hộ đã đọc.
3. Theo dõi số hộ thành công, lỗi và cảnh báo trong **Tiến trình và nhật ký**.
4. Sau khi xuất Excel thành công, bấm **Xuất PDF và tổng hợp** để chuyển các phiếu hợp lệ của đợt vừa xuất sang PDF và gộp theo nguồn.
5. Bấm **Mở thư mục kết quả** để xem các tệp đã tạo.

Tùy chọn **Kiểm tra khi xuất PDF** áp dụng cho bước chuyển PDF của đợt Excel vừa xuất:

- **Bỏ tích:** chuyển trực tiếp các tệp Excel sang PDF theo thiết lập in đã lưu, bỏ kiểm tra lại Excel và đối chiếu nội dung/bố cục PDF; thời gian xử lý thường ngắn hơn.
- **Tích:** kiểm tra lại Excel và đối chiếu PDF; thời gian xử lý dài hơn. Nếu đã sửa Excel sau khi xuất, bước kiểm tra có thể từ chối tệp đã thay đổi.

### Chuyển Excel có sẵn sang PDF

1. Chọn **Thư mục kết quả**.
2. Bấm **Excel có sẵn → PDF**.
3. Chọn một hoặc nhiều tệp `.xls`, `.xlsx`, `.xlsm` cần chuyển.
4. Chờ hoàn tất, xem nhật ký và mở thư mục kết quả.

Chức năng này không cần đọc dữ liệu nguồn hoặc chọn biểu mẫu. Ứng dụng dùng vùng in và thiết lập trang đã lưu trong từng tệp, không kiểm tra theo biểu mẫu PCGD. Hãy chỉnh vùng in, khổ giấy và hướng giấy trong Excel rồi lưu trước khi chuyển. Các tệp được nhóm theo thư mục chứa để tạo PDF tổng hợp.

### Dừng tác vụ

Bấm **Dừng xuất** khi đang xuất Excel/PDF. Ứng dụng sẽ dừng ở các điểm kiểm tra giữa các phần việc và đóng tài nguyên Excel; thao tác đang chạy có thể cần thêm thời gian để kết thúc. Chờ giao diện hết trạng thái bận trước khi chạy đợt mới. Các tệp đã xuất thành công vẫn được giữ lại; đợt bị dừng có thể chưa có PDF tổng hợp.

## 5. Vị trí kết quả

Ví dụ nguồn có tên `to_3.xls`, thư mục kết quả được chọn là `output`:

```text
output/
├── excel/
│   └── to_3/       # Phiếu Excel theo hộ
├── pdf/
│   └── to_3/       # PDF riêng theo hộ
└── to_3.pdf        # PDF tổng hợp của nguồn
```

Khi xử lý nhiều nguồn, mỗi nguồn có nhóm kết quả riêng. Nếu tên đã tồn tại, ứng dụng thêm hậu tố như `_2`, `_3` để tránh ghi đè. Xem nhật ký để biết đường dẫn chính xác của từng đợt.

## 6. Tốc độ và cấu hình

- **Luồng đọc (1–8):** số luồng đọc dữ liệu nguồn.
- **Xuất song song (1–4):** số tệp nguồn được xuất đồng thời; mỗi tiến trình dùng phiên Excel riêng. Tùy chọn này hữu ích khi có nhiều tệp nguồn; không chia một nguồn thành nhiều tiến trình.
- **Lưu cấu hình:** lưu lựa chọn vào `config/settings.json`. Khi chuyển máy, chọn lại đường dẫn nguồn, mẫu và kết quả vì đường dẫn đã lưu có thể thuộc máy cũ.
- `config/template_mapping.json` chứa ánh xạ ô trên biểu mẫu. Sao lưu trước khi chỉnh cấu trúc mẫu hoặc ánh xạ.

## 7. Xử lý lỗi thường gặp

| Hiện tượng | Cách xử lý |
| --- | --- |
| Không tìm thấy dữ liệu hoặc biểu mẫu | Chọn lại đường dẫn trên giao diện; bổ sung tệp Excel cục bộ nếu vừa tải mã từ GitHub. |
| Không đọc được dữ liệu hoặc thiếu hộ | Kiểm tra sheet, dòng tiêu đề, dòng bắt đầu và cột số phiếu; xem lỗi/cảnh báo trong nhật ký. |
| Không khởi tạo được Excel hoặc lỗi COM | Mở Microsoft Excel thủ công, hoàn tất các hộp thoại khởi động, rồi chạy ứng dụng trong phiên Windows đang đăng nhập. |
| Không ghi được kết quả | Chọn thư mục có quyền ghi và đóng tệp kết quả đang mở trong Excel hoặc trình đọc PDF. |
| PDF bị lệch trang khi chuyển Excel có sẵn | Mở tệp trong Excel, chỉnh vùng in và thiết lập trang, lưu rồi chuyển lại. |
| Xuất chậm khi có nhiều nguồn | Điều chỉnh **Xuất song song** phù hợp với máy; giảm số tiến trình nếu Excel quá tải. |

## 8. Kiểm thử dành cho phát triển

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements-dev.txt
.\venv\Scripts\python.exe -m unittest discover -s tests
```

Một số bài kiểm thử cần Excel COM hoạt động và các tệp mẫu/dữ liệu cục bộ đúng đường dẫn được khai báo trong bài kiểm thử. Các tệp này không có trong bản tải GitHub, nên cần chuẩn bị trước khi chạy toàn bộ bộ kiểm thử.

Tài liệu kỹ thuật và ghi chép dự án nằm trong `docs/` và [PROJECT_STATUS.md](PROJECT_STATUS.md).
