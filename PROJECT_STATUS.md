# PCGD_HOGIADINH — Trạng thái dự án

## Xuất nhiều nguồn song song — 2026-10-21

- GUI có mục **Xuất song song (1–4)**, mặc định 2 tiến trình, lưu trong cấu hình.
  Mỗi tiến trình dùng Excel COM riêng, xử lý từng nguồn; không chia sẻ COM.
- Hai nguồn giả lập, bỏ kiểm tra PDF: tổng thời gian giảm từ 36,74 xuống
  22,41 giây (khoảng 39%). Hash nguồn, mẫu và mapping giữ nguyên.
- Giữ thứ tự nguồn/hộ, tên thư mục và PDF tổng hợp; giữ chỗ tên trước khi chạy,
  chống ghi đè. Nguồn lỗi được báo riêng, nguồn khác tiếp tục.
- Chi tiết, giới hạn và lệnh benchmark: [PARALLEL_EXPORT.md](docs/PARALLEL_EXPORT.md).
  Hồi quy liên quan: **10 PASS, 0 FAIL, 0 SKIP**; chưa chạy lại toàn bộ suite.
  Khởi động lại GUI để dùng mã mới. Chưa commit thay đổi này.

## Tùy chọn kiểm tra khi xuất PDF — 2026-10-21

- Đã có checkbox **Kiểm tra khi xuất PDF**, mặc định bỏ tích, lưu bằng
  nút Lưu cấu hình. Khóa lựa chọn trong lúc tác vụ đang chạy.
- Bỏ tích: xuất trực tiếp Excel hiện tại, không nghiệm thu lại Excel/mẫu,
  không đối chiếu nội dung hoặc bố cục PDF. Giữ PageSetup đã lưu, ghi nhận
  số trang thực tế; vẫn đọc PDF để gộp, chống ghi đè và giải phóng COM.
- Tích: giữ kiểm tra đầy đủ như bản trước. API `export_pdf_batch` mặc định
  vẫn kiểm tra; GUI truyền rõ lựa chọn `validate=True/False` cho mọi nguồn.
- Bước xuất/nghiệm thu Excel không đổi. Không sửa nguồn, mẫu hoặc mapping.
  Commit mốc ổn định `594fe00` vẫn giữ nguyên; chưa commit các thay đổi mới.
- Kiểm thử liên quan: **10 test PASS, 0 FAIL** (9 test chức năng/cấu hình
  và 1 test COM có hai subtest bật/tắt). Có kiểm tra nhiều nguồn, 8 người
  trên 2 trang và tiếp tục sau lỗi COM ở cả hai chế độ. Chưa chạy toàn bộ suite.
- Đo một hộ giả lập: bước PDF bỏ kiểm tra mất **3,43 giây** so với khoảng
  15 giây ở chế độ đầy đủ; Excel+PDF mất 20,72 giây. Hash nguồn/mẫu/mapping
  trước/sau không đổi. Log: `docs/EXPORT_PDF_FAST_MODE.json`.

## Tối ưu tốc độ xuất — 2026-10-21

- Profiler xác định lặp gọi COM khi đọc đường viền là nút thắt lớn.
  `_com_cell_format` nay giữ đối tượng Borders/Border riêng trong từng ô,
  vẫn kiểm tra đủ thuộc tính, không cache qua workbook/luồng.
- Cùng một hộ giả lập: Excel+PDF giảm từ 43,43 xuống 33,42 giây (~23%);
  số lần gọi COM giảm từ 33.969 xuống 25.689 (~24%). Đây là benchmark một
  hộ trên máy hiện tại, chưa phải số đo 38 hộ hay cam kết cho mọi dữ liệu.
- Nhóm hồi quy sau tối ưu **11 PASS, 0 FAIL**; có test sửa đường viền vẫn
  bị chặn, font/phân trang, lỗi COM, tên nguồn, không ghi đè và GUI worker.
  Không chạy lại toàn bộ suite. Nguồn/mẫu/mapping không đổi, không commit.
- Chi tiết và lệnh đo: `docs/EXPORT_PERFORMANCE.md`. Khởi động lại GUI
  để nạp mã tối ưu. Không thêm các phiên Excel COM chạy song song.

## Cập nhật đường dẫn đầu ra theo nguồn — 2026-10-21

- GUI xuất vào `output/excel/<tên nguồn>/` và `output/pdf/<tên nguồn>/`;
  PDF tổng hợp là `output/<tên nguồn>.pdf`, không dùng `dot_<id>` cho đợt mới.
- Nhiều file input: to1.xlsx/to2.xlsx có nhóm to1/to2 độc lập và hai PDF
  tổng hợp to1.pdf/to2.pdf. Thư mục chỉ nạp một file dùng tên thư mục.
- Trùng tên hoặc xuất lại thêm `_2`, `_3`...; không ghi đè. Mã hộ trùng
  giữa các file được xuất riêng; trùng trong cùng file vẫn bị chặn.
- Giữ mapping, font cục bộ, phân trang, nghiệm thu PDF và tác vụ nền GUI.
  Cấu hình mẫu đã trỏ đúng vị trí hiện tại `tempplates/maubaocao.xls`.
- Các thư mục `dot_...` trong lịch sử bên dưới vẫn là kết quả cũ được giữ
  nguyên; không di chuyển hoặc đổi tên dữ liệu đã xuất của người dùng.
- Kiểm thử phần thay đổi: **10 PASS, 0 FAIL**, gồm Excel/PDF COM thật,
  tên nguồn/thư mục, trùng tên, xuất lại không ghi đè PDF tổng hợp, GUI
  worker và tiếp tục sau lỗi COM. Toàn bộ dữ liệu thử trong thư mục tạm;
  không tạo đợt thử trong output chính thức. Không chạy lại toàn bộ suite.

## Bản chính thức sau dọn dẹp — 2026-10-21

- Giữ mã nguồn, bộ test, công cụ phát triển, cấu hình/mapping, venv,
  tài liệu và các workbook/PDF đầu vào ở thư mục gốc.
- Kết quả chính thức mới nhất do người dùng chạy: 38 Excel tại
  `output/excel/dot_63c82a742cb749d58d28cf05cf5fdcbc/`; 38 PDF riêng và
  `TongHop.pdf` 38 trang tại `output/pdf/dot_599268997abf4ab78702846770207ff8/`.
- Đã xóa dữ liệu xuất thử, log chẩn đoán, ảnh render, cache, thư mục tạm
  và các đợt xuất cũ/trùng. Không chạy lại xuất dữ liệu khi dọn dẹp.
- Các đường dẫn test/chẩn đoán/PDF cũ trong ghi chép lịch sử bên dưới
  không còn tồn tại. Biên bản và hash file giữ nguyên: `docs/CLEANUP_20261021.json`.
- Nguồn, mẫu, cấu hình, mapping và toàn bộ kết quả chính thức được giữ
  nguyên hash trước/sau dọn dẹp. Không commit.

## Chẩn đoán PDF mới nhất — 2026-10-21

- Quy tắc hiện hành: tối đa **7 người/sheet**, mỗi sheet một trang A4 ngang,
  Zoom=False, FitToPagesWide/Tall=1, lề 0,5 cm. Thay thế 6 người/68%.
  Giữ đủ khối trống, đầu phiếu và vùng ký; giữ hình học/font mẫu.
- Đã chứng minh PDF cũ của TNg_QTRI_T18_12 có 2 trang thay vì 1 do Zoom
  68% vượt chiều cao in. Đếm trang Excel trước render bị cache. Sau sửa
  hộ này đúng 1 trang. Kiểm tra số trang, nội dung và chống sửa file vẫn giữ.
- Sửa thêm so sánh nội dung PDF khi ô xuống dòng bị xen văn bản và khi
  Excel General in số 6.0 thành 6. Có test RED trước sửa/GREEN sau sửa.
- Chạy ca biên 1/7/8/14/15: đúng 1/1/2/2/3 trang. Test một lần COM lỗi
  không ảnh hưởng hộ tiếp theo PASS.
- Sau khi người dùng duyệt giảm font cục bộ, ô lớp dài trên bản xuất dùng
  8 pt, ô lớp ngắn/trống giữ nguyên font mẫu. Hộ 110 đã PASS, hiển thị đủ
  `11 CHUYÊN TRUNG`, không đổi hàng/cột/vùng gộp hoặc workbook mẫu gốc.
- PDF hiện đủ **38 hộ, 38 trang, 0 lỗi**: giữ nguyên 37 PDF đã được duyệt,
  xuất lại riêng hộ 110 và đối chiếu lại nội dung cả 38 trước khi gộp.
  Kết quả: `output/pdf/dot_fbf45f4d5f844c89b53308f426b0f77c/TongHop.pdf`.
  Hash nguồn/mẫu/mapping trước/sau không đổi, không commit.
- Sau sửa font: nhóm hồi quy PDF/COM/PageSetup/nội dung **10 PASS, 0 FAIL**;
  chưa chạy lại toàn bộ suite, giữ riêng kết quả lịch sử bên dưới.
- Log, traceback, ảnh render và đường dẫn PDF: `docs/PDF_FAILURE_DIAGNOSIS.md`.
- Lượt toàn bộ trước thay đổi font: **59 test, 54 PASS, 5 FAIL, 0 ERROR, 0 SKIP**. Unittest
  ghi 6 failure events vì một method có 2 subtest FAIL. Bốn test còn kỳ vọng
  nguồn 427 người/120 hộ; một test kỳ vọng PrintArea mẫu rỗng đã lỗi thời.
  Các test PDF/COM hồi quy đều PASS. Log: `output/diagnostics/suite_20261021_162159/`.
- Khởi động lại GUI để nạp mã mới: `.\venv\Scripts\python.exe main.py`.
  Cần xuất lại Excel trước PDF; không dùng đợt cũ còn Zoom 68%.

## Cập nhật mới nhất — đọc thư mục, đọc song song và xuất chính thức

- GUI có hai lựa chọn Chọn tệp / Chọn thư mục nguồn. Thư mục được đọc
  đệ quy; bỏ mẫu, tệp tạm, thư mục kết quả và nhánh không thuộc dữ liệu.
- Có Luồng đọc 1–8, mặc định 4, lưu cùng cấu hình. Reader dùng pool có
  giới hạn, giữ thứ tự nguồn, khử trùng đường dẫn, báo lỗi từng tệp và
  cập nhật tiến độ. Excel COM vẫn tuần tự trên worker, không dùng chung
  COM giữa các luồng. Tra cứu kết quả hộ được tối ưu bằng dictionary.
- Nút Xuất thử 1 hộ đã thay bằng Xuất hộ đã chọn, đi qua cùng workflow
  chính thức như xuất tất cả. Excel vào `<đầu ra>/excel/dot_<id>/`, PDF
  vào `<đầu ra>/pdf/dot_<id>/`; không tự tạo thư mục test hoặc ghi đè.
  Xuất một hộ cũng tạo đợt Excel hợp lệ để tiếp tục xuất PDF trên GUI.
- Giữ mapping, 6 người/trang, 68%, Dân tộc 8 pt và checkbox đã duyệt.
  Đổi nguồn sau khi nạp sẽ yêu cầu đọc lại trước khi xuất.
- Đã kiểm tra cú pháp tĩnh; cập nhật/bổ sung test nhưng không chạy test,
  benchmark hoặc xuất dữ liệu theo yêu cầu trước của người dùng. Chưa
  tuyên bố kết quả PASS/FAIL mới hoặc mức tăng tốc thực tế. Không commit.
- Hướng dẫn hiện hành: `docs/OFFICIAL_GUI_WORKFLOW.md`.

## Cập nhật mới nhất 2026-10-21 — bàn giao mã để người dùng tự thử

- Theo xác nhận mới: chia 6 người/sheet, in A4 ngang ở Zoom 68%, lề
  0,5 cm; giữ sáu khối kể cả khối trống, đầu phiếu và vùng ký mỗi trang.
- Đã bỏ nhân độ rộng cột và đổi chiều cao dòng. Bản xuất lấy kích thước
  trực tiếp từ `maubaocao.xls` hiện tại, bao gồm các cột người dùng vừa sửa.
  Không đổi mapping, checkbox ☒/☐ tiếp tục dùng Segoe UI Symbol.
- Theo yêu cầu mới nhất, giảm phần Dân tộc (ô DT:, D12/D17/... theo
  mapping) xuống 8 pt trên bản xuất; giữ font và các phần khác của mẫu.
  Kiểm tra nghiệm thu font đã cập nhật tương ứng, không tắt bảo vệ định dạng.
- Đã dừng lượt test đang chạy theo yêu cầu người dùng; không chạy thêm
  test, không xuất thêm Excel/PDF sau thay đổi font này. Không công bố
  kết quả xanh cho phiên bản mới; người dùng tự chạy thử GUI.
- Trước khi đổi font đã có bộ giả lập 6 Excel/6 PDF, tổng hợp 15 trang:
  `output/testing/gui_batch_pdf_20261021_024156/`. Bộ này chưa chứa thay
  đổi font Dân tộc 8 pt và các sửa mẫu về sau của người dùng.
- Lượt toàn bộ trước đó: 51 test, 2 FAIL và 2 lỗi subtest COM; một FAIL
  do kỳ vọng phân trang cũ đã sửa và chạy riêng PASS; một FAIL do mẫu
  được lưu lại trong lúc chạy. Lỗi COM PrintArea rỗng đã sửa ở bản sao
  test, nhưng lượt chạy lại bị dừng trước khi hoàn tất. Giữ kiểm tra
  hash trước/sau và ngắt trang, không cập nhật snapshot để che lỗi.
- Chạy GUI: `.\venv\Scripts\python.exe main.py`. Khởi động lại GUI để
  nạp mã mới. Không sửa workbook gốc, không commit.

## Cập nhật 2026-10-21 — checkbox Unicode và chuẩn hóa A4

- Nữ/diện cư trú dùng ☒/☐ với Segoe UI Symbol; kiểm tra giá trị/font sau lưu.
  Không đổi ánh xạ các trường khác, nguồn, mẫu, mapping hoặc snapshot.
- Hai dạng phiếu 7/14 khối đã có cùng hình học vùng in trên bản xuất.
  A4 ngang, A1:AG87, lề 0,5 cm, header/footer 0, căn giữa ngang/dọc,
  Zoom=False và FitToPagesWide/Tall=1; Excel và PDF cùng một PageSetup.
- Bổ sung cảnh báo ngày sinh không hợp lệ và ô quá dài; chỉ vừa riêng ô,
  không đổi cỡ chữ toàn bảng. PDF thiếu nội dung trường đã ghi sẽ bị chặn.
- Toàn bộ test: **50 PASS, 0 FAIL, 0 SKIP, 0 ERROR**, 576,026 giây.
  Gồm COM Excel/PDF thật, checkbox, hình học/căn giữa, ngày sinh, GUI,
  chống ghi đè/tamper và các kiểm tra cấu trúc/mapping cũ.
- Báo cáo nguyên nhân, số đo và PDF bàn giao: `docs/PRINT_A4_CHECKBOX_REVIEW.md`.
  Bảng tận dụng 99,94% vùng in ngang; chữ tên thực tế khoảng 4,56 pt,
  vẫn nhỏ do giữ 14 khối × 5 dòng trong một trang A4.
- Chỉ xuất dữ liệu giả lập; không xuất hàng loạt dữ liệu thật, không commit.

## Cập nhật mới nhất 2026-10-21 — GUI Excel hàng loạt và PDF

- Đã nối nút Xuất tất cả Excel vào exporter hiện có qua workflow; nút mở
  sau khi nạp dữ liệu. Xuất toàn bộ hộ, mỗi đợt thư mục mới, không ghi đè.
- Nút Xuất PDF và tổng hợp mở khi đợt Excel gần nhất có phiếu hợp lệ.
  Nghiệm thu lại toàn bộ đợt, chuyển qua Excel COM, gộp theo mã hộ; chỉ
  PDF hợp lệ được gộp. Có tiến độ, nhật ký, thống kê, khóa tác vụ nền chung.
- Sửa vấn đề môi trường Excel tự đổi A4 sang Letter bằng MapPaperSize=False
  trong phiên PDF riêng, khôi phục khi kết thúc. Không đổi mapping hoặc mẫu.
- Bộ thử mới: `output/testing/gui_batch_pdf_20261021_014204/`, 5 hộ tổng
  hợp 1/7/8/14/15 người: 5 Excel, 5 PDF, 0 lỗi, tổng hợp 6 trang. Render
  đủ 6 trang đã xem; nguồn/mẫu/mapping trước/sau không đổi.
- Toàn bộ test: **48 PASS, 0 FAIL, 0 SKIP** (431,080 giây). Có test GUI,
  COM PDF thật, thứ tự gộp, chống ghi đè, mã trùng/thiếu và Excel bị sửa
  sau nghiệm thu. Kiểm tra metadata mã dự phòng riêng cũng PASS.
- Hướng dẫn và giới hạn: `docs/GUI_BATCH_PDF.md`. PDF dùng đợt Excel gần
  nhất trong phiên GUI; chưa có màn hình chọn lại thư mục Excel sau restart.
  Giữ nguyên ô mẫu nên cần rà soát chuỗi quá dài ở cột hẹp.
- Không chạy xuất hộ thật trong lượt này; không commit hoặc tự tuyên bố
  nghiệm thu nghiệp vụ.

## Cập nhật mới nhất 2026-10-21 — cấu trúc mẫu và đối chiếu 6 file

- Theo xác nhận người dùng, SHA-256 mẫu chỉ ghi log; không chặn theo hash,
  mtime, kích thước/path snapshot cũ. Giữ nguyên snapshot lịch sử.
- Thay test hash bằng test COM cấu trúc thực tế; tăng kiểm tra toàn bộ
  14 khối, 70 nhãn năm, vùng ký, A4 ngang và bảo vệ công thức/external links.
- Kiểm tra trực tiếp 6 file cũ tại `approved_mapping_20261021_010939`:
  giá trị và phân trang đúng; phát hiện NumberFormat Z2 ở hộ 5/7/14 là Text
  thay vì General. Giữ nguyên file cũ; sửa exporter khôi phục định dạng Z2.
- Bộ mới tại `output/testing/approved_mapping_20261021_012426/` đã đối
  chiếu 20.097 ô/7 sheet: 0 sai lệch; <=14 người 1 trang, 15 người 2 sheet
  x 1 trang. Nguồn/mẫu/snapshot và 6 file cũ không đổi.
- Toàn bộ test cuối: **43 PASS, 0 FAIL, 0 SKIP** (316,940 giây).
  Báo cáo: `docs/TEMPLATE_STRUCTURE_AND_SAMPLES_REVIEW.md` và hai audit JSON.
- Không đổi mapping/GUI, không commit, không xuất hộ thật hoặc mở khóa
  Xuất tất cả/PDF. Không tự tuyên bố nghiệm thu; dừng chờ kiểm tra GUI.

## Cập nhật 2026-10-21 — mapping theo phê duyệt mới

- Chi tiết hiện hành: `docs/APPROVED_MAPPING_20261021.md`, thay thế các ghi
  chú cũ tương ứng về mapping chưa duyệt, nhãn năm linh hoạt và phiếu ngắn.
- Đã bổ sung Xóm A3, cư trú, Nữ, TN nghề, học xong/bỏ học, xóa mù chữ,
  khuyết tật/CNKT/KNHT, chi tiết hoàn cảnh và ghi chú. Các trường theo tên
  dùng tiêu đề thực, cảnh báo và không ghi nếu thiếu/trùng tiêu đề.
- Nhãn cố định 2025-2026 đến 2029-2030; dữ liệu chỉ trên dòng năm thứ nhất.
  Hộ 1–7 hiện 7 khối, 8–14 hiện 14 khối; giữ sheet bổ sung trên 14.
- Xuất sáu hộ tổng hợp 1/5/7/8/14/15 người tại
  `output/testing/approved_mapping_20261021_010939/`; Excel xác nhận hộ
  <=14 in 1 trang, hộ 15 in 2 sheet x 1 trang. Hash nguồn/mẫu/snapshot
  trước và sau xuất thử không đổi.
- Bộ test cuối: **42 PASS, 1 FAIL, 0 SKIP** (43 test). FAIL duy nhất là
  hash `maubaocao.xls` hiện tại `fd4211e3...` khác snapshot `963a75a4...`.
  Không sửa snapshot hoặc bỏ kiểm tra bảo vệ. Mapping và test COM đều PASS.
- Giữ nút Xuất tất cả/PDF vô hiệu hóa. Không commit, không xuất hàng loạt,
  không tuyên bố nghiệm thu; chờ người dùng kiểm tra GUI và phiên bản mẫu.

Cập nhật: 2026-10-20

## Giai đoạn hiện tại

**Giai đoạn 3 — Xuất Excel theo hộ: exporter đã triển khai và kiểm thử bằng
Excel COM thật; ánh xạ nghiệp vụ chưa hoàn tất.**

Giai đoạn 1 và 2 đã hoàn tất. Giai đoạn 3A có kiểm thử tổng hợp cho hộ
1/6/7/14/15 người trên mẫu thật. Chưa cho phép coi toàn bộ nghiệp vụ Excel
đã nghiệm thu vì các ánh xạ chưa chốt và địa bàn chưa được xác nhận. Năm học
`2026-2027` đã cấu hình; người dùng đã duyệt Phương án A cho nhãn linh hoạt,
điểm neo là nhãn đầu tiên và đã kiểm thử trên bản sao. Mapping B5–B12 vẫn
chưa được duyệt. GUI Tkinter đang được xây dựng để đọc dữ liệu, xem hộ và
xuất thử một hộ; PDF và xuất hàng loạt vẫn chưa triển khai.

Đã có phê duyệt thử mapping B5–B7, nhưng chưa nghiệm thu nghiệp vụ. B5–B7
ghi tại các ô neo của dòng đầu theo cấu hình ánh xạ; B8–B12 và các trường
khác chưa duyệt tiếp tục không được ghi.

Đã tạo:

- `docs/DATA_ANALYSIS.md` — sheet/cột nguồn, thống kê dữ liệu mẫu và cảnh báo.
- `docs/TEMPLATE_MAPPING.md` — cấu trúc phiếu, vùng gộp, hàng thành viên,
  ánh xạ và thử nghiệm phiếu ngắn.
- `docs/MAPPING_AUDIT.md` — giải thích 15 lần cảnh báo kiểm thử, phân loại
  đủ 36 trường chưa ánh xạ A/B/C, bằng chứng đọc mẫu và điểm cần xác nhận.
- `docs/PHASE3_FINAL_AUDIT.md` — kết quả audit Giai đoạn 3, độ phủ mapping,
  dữ liệu thật, cấu hình và cổng nghiệm thu còn chặn.
- `docs/PHASE3_DECISION_CHECKLIST.md` — metadata/hash nguồn và các câu hỏi
  cần phê duyệt trước khi xuất thật.
- `docs/USER_DECISIONS.md` — đầy đủ 15 nhóm câu hỏi B1–B15, trường/cột
  nguồn, ô đích dự kiến và quy tắc chờ người dùng phê duyệt.
- `docs/SCHOOL_YEAR_MAPPING_AUDIT.md` — bằng chứng COM cho 70 ô nhãn năm,
  các vùng năm lịch, điểm neo còn mở và phân tích B5–B12.
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
  xuất báo cáo XLSX, tải cấu hình ngoài; `save_settings(...)` cập nhật an
  toàn xã/năm học để GUI tương lai sử dụng.
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
- `config/settings.json` — cột nguồn, tên sheet, năm học đã cấu hình và
  xã/phường còn chờ người dùng cung cấp; không hard-code địa bàn.
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
- Trong lượt kiểm tra hiện tại, hash trước/sau khi đọc nguồn và mở lại
  workbook không đổi. Hash hiện tại của nguồn khác hash trong báo cáo đối
  chiếu cũ; báo cáo cũ không ghi path/size/mtime, chưa xác định nguyên
  nhân. Hiện tại: `dulieumau.xls`, 623104 bytes,
  `C:\Users\Admin\Desktop\VSCODE PROJECT\PHOCAPGD\dulieumau.xls`,
  mtime 2026-10-19 14:25:01.044 UTC, SHA-256
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
- Bản thử có thể mở lại tại `output/excel/kiem_thu_giai_doan_3/`; báo cáo
  đối chiếu đặt cạnh các workbook. Đây là dữ liệu tổng hợp, không phải phiếu
  hộ thực tế.

## Lộ trình cập nhật theo yêu cầu Excel/PDF

| Chặng | Phạm vi | Trạng thái |
|---|---|---|
| 1. Phân tích dữ liệu/mẫu | Đọc workbook, cấu trúc và thử bố cục phiếu ngắn | Hoàn tất |
| 2. Reader và nhóm hộ | Đọc `.xls`/`.xlsx`, giữ dữ liệu thô, chuẩn hóa khóa, cảnh báo trùng/xung đột và báo cáo | Hoàn tất; 14/14 test PASS |
| 3A. Xuất Excel | Một `.xlsx`/hộ; phiếu ngắn/đầy đủ/bổ sung | Đã triển khai; test tổng hợp 1/6/7/14/15 PASS |
| 3B. Nghiệm thu Excel | Mở lại file, xác minh thành viên, đánh số, cấu trúc, merge, vùng in, xác nhận và định dạng | 21 PASS, 0 FAIL, 0 SKIP; chỉ fixture tổng hợp, nghiệm thu nghiệp vụ còn chặn do cấu hình và mapping chưa chốt |
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
`.\venv\Scripts\python.exe -m unittest discover -v` — **21 PASS, 0 FAIL,
0 SKIP**.
Mười chín test trước tiếp tục PASS; thêm hai test preflight cấu hình. Các
test Excel dùng Microsoft Excel COM thật trên Windows và `maubaocao.xls`
chỉ đọc.

Đã tạo 5 workbook kiểm thử tổng hợp cho 1, 6, 7, 14 và 15 thành viên cùng
báo cáo đối chiếu. Báo cáo xác nhận số người/thứ tự, 442 vùng gộp mỗi sheet,
phần xác nhận, vùng in, external links và hash trước/sau của lượt test khi
đó. Hash nguồn hiện tại khác giá trị ghi trong báo cáo cũ.
Chi tiết nằm trong `output/excel/kiem_thu_giai_doan_3/BAO_CAO_DOI_CHIEU.json`.

Đã audit 15 lần cảnh báo ở 5 workbook: mỗi workbook có cảnh báo diện cư trú,
năm học chưa ánh xạ và giới tính chưa ánh xạ. Đã thêm diện ưu tiên J vào
danh sách cảnh báo trường chưa ánh xạ. Không tìm thấy trường mới đủ căn cứ
để ghi; tổng mapping hiện có 10 quy tắc đích/13 cột nguồn, với 36 trường
nghiệp vụ còn khai báo chưa ánh xạ.

Kiểm tra Excel COM trực tiếp chỉ đọc mẫu xác nhận: `C12` có placeholder
`q`/Wingdings cạnh nhãn Nữ; `Z2` in lựa chọn `q Thường trú`/`q Tạm trú`;
`F9:F13` in các năm 2025–2030, còn các cột dữ liệu theo năm là ô nhập
trống. Các vị trí này chưa đủ xác nhận ký hiệu chọn hoặc luật gắn năm nên
không thêm mapping. Phân loại 36 trường: A=0, B=31, C=5; xem
`docs/MAPPING_AUDIT.md`.

Lượt xác thực cuối ngày 2026-10-20: `compileall` thành công;
`config/settings.json`, `config/template_mapping.json` và báo cáo JSON parse
được. `src/settings.py` không tồn tại trong cấu trúc hiện tại; module cấu
hình thực tế là `settings.py` ở thư mục gốc.

Đã tạo `docs/PHASE3_DECISION_CHECKLIST.md`, gồm những gì đã xác minh/chưa
xác minh, quyết định chặn xuất thật, hạng mục có thể hoãn, so sánh hash và
tiêu chí đề xuất nghiệm thu. Bảng 31 trường nhóm B và câu hỏi B1–B15 nằm
trong `docs/MAPPING_AUDIT.md`. Không bổ sung mapping, không xuất hộ thật,
không chạy lại suite vì lượt này chỉ thay đổi tài liệu.

## Việc chưa xử lý / cần xác nhận

1. `config/settings.json` hiện có `commune_name=null`,
   `school_year="2026-2027"`. `settings.py` nạp được cấu hình; hàm
   `save_settings(...)` cập nhật an toàn hai khóa mà không làm mất cấu hình
   khác. Exporter chặn xã/năm học thiếu, năm học sai định dạng hoặc địa bàn
   hộ không khớp cấu hình trước khi tạo thư mục đầu ra/khởi chạy COM.
2. Hash `dulieumau.xls` hiện tại không khớp báo cáo đối chiếu cũ; chưa có
   bản nguồn cũ/metadata đủ để xác định nguyên nhân. Người dùng cần xem xét
   và phê duyệt phiên bản nguồn trước khi xuất thật.
3. Người dùng duyệt năm cấu hình ở nhãn đầu tiên, giữ nguyên năm lịch ký và
   cho phép bản sao thử. Hàm riêng cập nhật 70 nhãn trên workbook bản sao;
   không nối vào luồng xuất hộ. B5–B12 không được ghi và vẫn cần duyệt.
4. Chưa ghi các trường J, M/O/G, P, Q:T, V:AH, AI:AR, AS:AX (ngoại trừ AU,
   AV, AW) khi chưa có quy tắc xác nhận/ô đích. Báo cáo exporter liệt kê
   cảnh báo cho các trường nguồn có dữ liệu.
5. Ký tự `q` trong checkbox là font Wingdings; ký tự thay thế để đánh dấu
   giới tính/diện cư trú chưa được xác minh. Các checkbox được giữ nguyên.
6. Kết quả xuất thử hiện dùng dữ liệu tổng hợp: 5 thành công, 0 thất bại.
   Chưa xuất `.xlsx` hộ thật cho 120 hộ vì xã chưa được xác nhận; dữ liệu
   nguồn không thay đổi.
7. PDF, gộp PDF và EXE chưa triển khai; không có test nào của các chặng đó
   được coi là PASS. GUI đang được triển khai; chưa được xem là nghiệm thu
   ứng dụng hoặc Giai đoạn 3.

## Bước tiếp theo

Tiếp theo cần tên xã thực tế và trả lời các quy tắc còn mở ở
`docs/USER_DECISIONS.md`/`docs/MAPPING_AUDIT.md`.
Sau đó bổ sung test cho từng mapping, xuất thử các
hộ nguồn đại diện (bao gồm hộ mâu thuẫn), đối chiếu Excel từng trường và
hoàn tất cổng nghiệm thu Excel. Chỉ sau khi cổng Excel nghiệp vụ PASS mới
bắt đầu Excel COM → PDF, xác thực PDF, gộp PDF và kiểm thử chạy lại từ Excel
sẵn có.

## Cập nhật 2026-10-20 — cấu hình và quyết định người dùng

- Đặt `school_year` thành `2026-2027`; giữ `commune_name` là JSON `null`
  do chưa được người dùng cung cấp. Không thay đổi các khóa cấu hình khác.
- Bổ sung `settings.save_settings(...)`: cập nhật xã/năm học bằng thao tác
  ghi tệp tạm và thay thế nguyên tử; giữ nguyên cấu hình khác. Xã rỗng/null
  vẫn bị validator từ chối trước khi xuất chính thức.
- Tạo `docs/USER_DECISIONS.md` với đủ 15 nhóm B1–B15. Không thêm mapping.
- Ghi nhận SHA-256 nguồn hiện tại:
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
  Trạng thái nguồn vẫn **CHỜ NGƯỜI DÙNG PHÊ DUYỆT**; chưa xuất hộ thật,
  chưa nghiệm thu Giai đoạn 3 và chưa commit.
- Xác thực lượt này: `unittest discover` **22 PASS, 0 FAIL, 0 SKIP**;
  `compileall` PASS; hai JSON cấu hình parse hợp lệ và giá trị yêu cầu đúng;
  `git diff --check` PASS.
- Tính lại SHA-256 nguồn sau kiểm thử: vẫn
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
  Không xuất hộ thật; thay đổi tài liệu/mã chưa commit.

## Cập nhật 2026-10-20 — audit năm học linh hoạt

- Người dùng duyệt **Phương án A – Năm học linh hoạt**, neo năm cấu hình
  vào nhãn đầu; các năm lịch ký giữ nguyên; cho phép bản sao thử. Không suy
  ra mapping dữ liệu giáo dục từ quyết định cập nhật nhãn.
- Excel COM read-only tìm thấy đúng 70 ô nhãn năm học tại cột F, mỗi ô gộp
  với G, thuộc 14 vị trí thành viên. Mỗi vị trí có cùng dãy 5 năm
  2025–2026 đến 2029–2030; chiều cao dòng/đường viền có khác biệt theo vị
  trí. Phần ký có các năm lịch riêng 2025–2029.
- Bổ sung `src/school_year_labels.py` với danh sách allowlist tường minh 70
  ô và kiểm tra vùng gộp/không đổi ô ngoài vùng cho phép. Hàm chưa nối vào
  luồng xuất hộ chính thức.
- Tạo hai workbook thử trong `output/testing/school_year_labels/` cho các
  năm `2026-2027` và `2027-2028`; Excel COM mở lại xác nhận 70 nhãn đúng,
  442 vùng gộp và cấu trúc mẫu được bảo toàn theo kiểm thử.
- Tạo `docs/SCHOOL_YEAR_MAPPING_AUDIT.md`; cập nhật `docs/SPEC.md` và
  `docs/USER_DECISIONS.md`, không đổi `config/template_mapping.json`, không
  ghi dữ liệu thành viên.
- Hash mẫu trước/sau không đổi:
  `89354830a91694e6a9251ccd0b02214c19625c211da683e1a086a74411694aad`.
  Hash nguồn trước/sau không đổi:
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
- Bổ sung test cho năm hợp lệ 2026–2027/2027–2028, định dạng sai và cặp
  không liên tiếp. Test xuất tổng hợp xác nhận các trường B5–B12 chưa duyệt
  không bị ghi vào ô đích ứng viên.
- Xác thực gần nhất trước lượt này: `unittest discover` **23 PASS, 0 FAIL,
  0 SKIP**. Kết quả sau triển khai được ghi dưới đây.

## Cập nhật 2026-10-20 — triển khai và thử nhãn năm

- Cập nhật theo điểm neo đã duyệt: nhãn đầu tiên của mỗi khối bằng năm
  cấu hình; 4 nhãn tiếp theo tăng liên tục; không dùng ngày máy tính.
- Tạo module riêng `src/school_year_labels.py`; chỉ cập nhật allowlist 70
  địa chỉ đã đối chiếu, xác minh mỗi ô là neo của merge `F:G`, kiểm tra
  không có nhãn thừa/thiếu, bảo toàn các năm ký và xác nhận chỉ 70 giá trị
  thay đổi. Luồng xuất hộ không bị sửa để gọi module này.
- Hai workbook thử nghiệm đã lưu dưới `output/testing/school_year_labels/`;
  không có dữ liệu hộ thành viên. Excel COM mở lại được, đúng nhãn 2026–
  2031 và 2027–2032, 442 merge; kiểm tra font/viền nhãn, chiều cao dòng,
  chiều rộng cột, trang/in, PrintArea, ngắt trang, vùng ký và external links
  trong test tích hợp.
- Mapping nghiệp vụ B5–B12 vẫn chưa được duyệt; không xuất 120 hộ thật.
- Kiểm thử cuối: `unittest discover` **28 PASS, 0 FAIL, 0 SKIP**. Trong
  module nhãn năm có 3 test unit và 2 test tích hợp Excel COM; bộ suite
  cũng chạy các test Excel COM của exporter. `compileall` PASS; JSON hợp lệ;
  `git diff --check` PASS.
- Hash trước/sau không đổi: mẫu
  `89354830a91694e6a9251ccd0b02214c19625c211da683e1a086a74411694aad`;
  nguồn
  `5382ace88e435e2c6d2746dc779ed700687c18dee4bce50c7b38cec5eca681c0`.
  Hash nguồn vẫn **CHỜ NGƯỜI DÙNG PHÊ DUYỆT** cho xuất thật; không commit.

## Cập nhật — xây dựng giao diện Tkinter

- Thêm `gui.py` và điểm khởi chạy `main.py`. GUI tiếng Việt có chọn tệp
  nguồn/mẫu/thư mục đầu ra, xã/phường, năm học, lưu cấu hình, phân tích nền,
  thống kê, tìm kiếm hộ, xem thành viên/cảnh báo và cửa sổ chỉ xem các
  mapping chưa được duyệt.
- Chỉ cho xuất thử đúng một hộ vào `<thư_mục_kết_quả>/test_output`; yêu cầu
  xã/phường, năm học hợp lệ, hộ đã được đọc với cùng xã/phường và xác nhận
  người dùng. Thông báo và log nêu rõ tệp thử chưa phải phiếu hoàn chỉnh.
  Exporter vẫn là nơi thực hiện Excel COM và xử lý đóng workbook. Luồng GUI
  bật cập nhật 70 nhãn năm trên bản sao; mặc định exporter cho luồng khác
  không đổi.
- Đường dẫn nguồn, mẫu, thư mục đầu ra cùng xã/năm học được lưu qua
  `settings.py`; cấu hình cũ vẫn được giữ nguyên. Không tự điền xã.
- Nút xuất tất cả Excel và xuất PDF bị vô hiệu hóa, hiển thị “Chưa sẵn sàng”.
  Không khởi chạy Excel khi mở GUI. Khi đang chạy tác vụ, không cho đóng
  cửa sổ để tránh bỏ dở Excel COM.
- Cách chạy: `.\venv\Scripts\python.exe main.py`.
- Xác minh GUI khởi động và chạy liên tục trong smoke test; tiến trình thử
  đã được đóng sau đó. `unittest discover -v`: **35 PASS, 1 FAIL, 0 SKIP**
  (36 test). 6 test GUI và test COM cập nhật nhãn năm cho bản sao đều PASS;
  thử đọc GUI trực tiếp nhận 427 nhân khẩu và 120 hộ. Một test tích hợp
  nhãn năm thất bại do hash mẫu hiện tại không khớp hash kỳ vọng đã ghim;
  không thay đổi giá trị hash test. `compileall`, kiểm tra hai JSON và
  `git diff --check` đều PASS.
- Hash đo được trước/sau kiểm thử không đổi trong lượt này:
  `dulieumau.xls` = `e7ce0b359bc516d54463a29ab88142ff355b735bda6408f77566dc5e0d3b262a`;
  `maubaocao.xls` = `963a75a4ae9258ff40f45006f97a9750c31023e0cee7bb4c115d938c52a9434e`.
  Cả hai khác hash từng được ghi nhận trước đây (nguồn `5382ace8...`,
  mẫu `89354830...`); chưa xác định nguyên nhân và không thay đổi/khôi phục
  hai tệp. Nguồn tiếp tục **CHỜ NGƯỜI DÙNG PHÊ DUYỆT**; không xuất hộ thật,
  không nghiệm thu Giai đoạn 3 và không commit.

## Cập nhật 2026-10-20 — snapshot workbook được người dùng xác nhận

- Người dùng xác nhận sử dụng phiên bản hiện tại của `dulieumau.xls` và
  `maubaocao.xls`. Lưu snapshot đường dẫn, kích thước, mtime UTC và SHA-256
  tại `config/approved_workbook_snapshot.json`; không sửa hai workbook.
- Tách test integrity `test_approved_workbook_snapshot_matches_current_files`
  khỏi test COM chức năng `test_updates_only_verified_labels_and_preserves_workbook_structure`.
  Test COM không còn phụ thuộc hash; vẫn kiểm tra 70 nhãn, ô ngoài phạm vi,
  442 vùng gộp, định dạng, kích thước hàng/cột, trang in, vùng ký và links.
- Hai test liên quan chạy riêng: **2 PASS, 0 FAIL**. Toàn bộ suite:
  **36 PASS, 1 FAIL, 0 SKIP**. Lỗi còn lại là test cấu hình kỳ vọng xã rỗng
  nhưng `config/settings.json` hiện có `commune_name="PDP"`; không thay đổi
  cấu hình địa phương để làm test PASS.
- `dulieumau.xls` SHA-256
  `e7ce0b359bc516d54463a29ab88142ff355b735bda6408f77566dc5e0d3b262a`;
  `maubaocao.xls` SHA-256
  `963a75a4ae9258ff40f45006f97a9750c31023e0cee7bb4c115d938c52a9434e`.
  Hash sau kiểm thử không đổi. JSON và `git diff --check` PASS. Không thêm
  mapping, không xuất hàng loạt, không commit.

## Cập nhật 2026-10-20 — thử mapping B5–B7

- Đã xác minh read-only bằng Excel COM 14 bộ ô đích trên mẫu: H:J (lớp),
  K:O (trường), P (bậc tốt nghiệp), R (năm tốt nghiệp), tại đúng các hàng
  đầu khối `9, 14, 19, 24, 29, 34, 39, 47, 52, 57, 62, 67, 72, 77`.
  Mapping được khai báo bằng danh sách địa chỉ và merge-area tường minh,
  không tính dòng bằng công thức.
- Triển khai B5: ghi nguyên R, Q chỉ đối chiếu; B6: ghi nguyên T, không ghi
  S; B7: ghi nguyên V/X vào P/R, kể cả năm cũ; W không ghi. Giá trị bất
  thường, không so sánh được/mâu thuẫn và cặp V/X thiếu một trường có cảnh
  báo. Dữ liệu chuỗi được ghi với định dạng văn bản tạm thời rồi khôi phục
  định dạng ô để Excel COM không tự chuyển chuỗi năm đơn sang số.
- Chưa nối việc cập nhật nhãn năm vào lượt xuất này theo lựa chọn người
  dùng: giữ nguyên cả 70 nhãn như mẫu. Mẫu bắt đầu ở `2025-2026`, cấu hình
  là `2026-2027`; vì vậy workbook thử không đồng bộ nhãn cấu hình và chưa
  thể dùng cho xuất chính thức.
- Dữ liệu chỉ tổng hợp giả lập. Tạo đúng ba workbook thử:
  `output/testing/education_mapping/ho_SYNTHETIC-06.xlsx`,
  `ho_SYNTHETIC-07.xlsx`, `ho_SYNTHETIC-14.xlsx`; bao phủ thành viên thứ
  6/7/14, lớp/trường, tốt nghiệp MN quá khứ, tốt nghiệp THPT quá khứ, không
  có dữ liệu học tập, Q/R bất đồng, T dạng năm và V/X thiếu cặp.
- Giai đoạn 3 chưa nghiệm thu; không xuất 120 hộ thật, không tạo PDF hàng
  loạt, không triển khai GUI và không commit.
- Lưu ý provenance: SHA-256 hiện quan sát được của `dulieumau.xls` là
  `e4d32c09c2d8e191f59e3dc991836a4aec13ee36de20dbb68369d1730dcf789f`,
  khác digest `5382ace8...` ghi ở audit trước. Tệp chỉ được hash ở lượt
  này, không được mở để đọc dữ liệu hay ghi lại; nguyên nhân khác biệt
  chưa xác định. Test cũ đang ghim digest trước nên kết quả suite phải được
  báo riêng, không tự đổi hash kỳ vọng hoặc coi hash mới đã được phê duyệt.
