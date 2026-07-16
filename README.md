# Ứng dụng Lọc trùng Danh sách Bệnh nhân THA

Ứng dụng desktop offline (Python + SQLite + giao diện PyQt6) để nhập danh sách
bệnh nhân từ Excel, lưu trữ, lọc trùng và xuất kết quả ra Excel/CSV.

Chỉ có **1 bộ cài đặt duy nhất** — khi cài, trình cài đặt sẽ hỏi vai trò của
máy này:
- **Một máy** — dùng độc lập, không chia sẻ qua mạng (mặc định).
- **Máy trạm** — kết nối tới 1 máy chủ khác đã có sẵn trong mạng LAN nội bộ.
- **Máy chủ** — chia sẻ dữ liệu của máy này cho các máy khác trong mạng LAN
  (máy này vẫn dùng được giao diện chính bình thường, đồng thời có thêm 1
  Windows Service chạy ngầm để các máy trạm kết nối vào). Xem mục "Chia sẻ
  dữ liệu cho nhiều máy" bên dưới.

Mã nguồn gồm:
- `core.py` — tầng dữ liệu: SQLite, đọc/chuẩn hóa Excel, xuất Excel/CSV (không phụ thuộc giao diện).
- `app.py` — giao diện PyQt6 (nhập `core.py` để xử lý dữ liệu).
- `netserver.py` / `netclient.py` — tầng giao tiếp qua mạng LAN dùng chung bởi
  cả máy chủ và máy trạm, chỉ dùng thư viện chuẩn Python.
- `service.py` / `server_tray.py` — thành phần chạy ngầm khi máy này được
  chọn vai trò "Máy chủ" lúc cài đặt (Windows Service + tray helper), xem
  mục "Chia sẻ dữ liệu cho nhiều máy" bên dưới. Được đóng gói cùng 1 file
  cài đặt với ứng dụng chính, không phải bộ cài riêng.

## Yêu cầu

- Python 3.8 trở lên (máy này đang có Python 3.10).
- Thư viện `PyQt6`, `PyQt6-Charts` (vẽ biểu đồ ở tab "Thống kê"), `openpyxl`
  (đọc/ghi `.xlsx`) và `xlrd` (đọc file `.xls` cũ định dạng Excel 97-2003) —
  đã có sẵn trên máy này. Nếu máy khác chưa có, chạy:
  ```
  pip install -r requirements.txt
  ```
- `sqlite3` đã có sẵn trong Python chuẩn, không cần cài thêm.

## Chạy ứng dụng

Cách 1: Bấm đúp vào `run.bat`.

Cách 2: Mở terminal tại thư mục này rồi chạy:
```
python app.py
```

Dữ liệu được lưu trong file `benh_nhan.db` (SQLite) ngay trong thư mục này —
hoàn toàn offline, không gửi dữ liệu ra ngoài.

## Các chức năng

Ở góc trên cùng cửa sổ có thanh menu **"Trợ giúp"**: **Kiểm tra cập nhật
ngay** (hỏi GitHub Releases ngay lập tức, không chờ kiểm tra ngầm lúc mở
app), **Chạy cập nhật (update.bat)** (mở luôn cửa sổ cập nhật — tự xin quyền
Administrator nếu máy đang ở vai trò Máy chủ), **Hướng dẫn sử dụng (README)**
và **Giới thiệu** (xem phiên bản hiện tại).

### 1. Tab "Nhập dữ liệu"
- Chọn file Excel (`.xlsx` hoặc `.xls` cũ) và bấm **Nhập vào cơ sở dữ liệu** —
  mở ra hộp thoại **"Ánh xạ cột khi nhập Excel"**:
  - Ứng dụng tự đoán trước: sheet nào, dòng tiêu đề nào, và cột nào ứng với
    trường nào (Họ và tên, Giới tính, Năm sinh, CCCD, Địa chỉ, Ngày khám, Chẩn
    đoán...) — dựa theo **tên cột thực tế** trong file (nhận diện được nhiều
    biến thể như "Tên bệnh nhân" thay vì "Họ và tên", "CCCD/ Hộ chiếu" thay vì
    "Số CCCD"...), không còn giả định thứ tự cột cố định như trước.
  - Bảng xem trước 5 dòng dữ liệu đầu giúp kiểm tra nhanh việc đoán có đúng
    không — sai/thiếu ô nào thì sửa lại bằng dropdown tương ứng trước khi bấm
    **"Nhập dữ liệu từ sheet này"**. Cột bắt buộc phải chọn: **Họ và tên**.
  - Một số file có **Giới tính tách thành 2 cột con "Nam"/"Nữ"** (giá trị năm
    sinh/tuổi nằm ở cột nào xác định giới tính đó) thay vì 1 cột "Giới tính"
    dạng chữ — tích ô **"Giới tính & Năm sinh/Tuổi nằm trong 2 cột tách
    riêng"** (tự bật sẵn nếu nhận diện được) rồi chọn đúng 2 cột.
  - File có **nhiều sheet bố cục khác nhau** (thường gặp ở báo cáo BKLN thực
    tế — mỗi mặt bệnh 1 sheet riêng, có khi kèm cả sheet tổng hợp số liệu
    không phải danh sách bệnh nhân): đổi **Sheet** ở đầu hộp thoại và lặp lại
    các bước trên cho từng sheet cần nhập — không bắt buộc nhập hết mọi sheet,
    sheet nào không phải danh sách bệnh nhân thì bỏ qua. Bấm **"Đóng"** khi
    nhập xong tất cả các sheet cần thiết.
- Mỗi dòng dữ liệu được nhận diện trùng lặp tuyệt đối (giống hệt mọi trường)
  bằng mã băm — nên **nhập lại cùng một file nhiều lần sẽ không bị nhân đôi
  dữ liệu**. Việc này khác với "bệnh nhân trùng" (cùng một người có nhiều lượt
  khám) — trường hợp đó xử lý ở tab "Lọc trùng".
- Có thể nhập nhiều file khác nhau, dữ liệu sẽ được gộp vào cùng CSDL.
- Ứng dụng tự động phát hiện và sửa các dòng bị đảo nhầm 2 cột "Giới tính" và
  "Ngày sinh" ngay khi nhập (ví dụ Giới tính ghi "1958" còn Ngày sinh ghi "Nam") —
  lỗi này có trong file Excel nguồn.
- Sau khi nhập, ứng dụng hiện **báo cáo chất lượng dữ liệu**: đếm số dòng thiếu
  CCCD/BHYT/chẩn đoán, không xác định được năm sinh hoặc ngày khám. Bấm
  **Xuất báo cáo chất lượng dữ liệu** để xem chi tiết từng dòng (kèm cột
  "Loại lỗi") ra Excel/CSV.
- Nút **Xóa toàn bộ dữ liệu trong CSDL** dùng khi muốn làm sạch để nhập lại từ đầu.
- Nút **Sửa lỗi đảo cột Giới tính / Ngày sinh** dùng để quét và sửa lại các dòng
  bị lỗi này trong dữ liệu đã nhập từ trước (trước khi có tính năng tự sửa khi nhập).
- Nút **Sửa lỗi định dạng Ngày khám bị bỏ sót** dùng khi Ngày khám ghi theo kiểu
  "HH:MM dd/mm/yyyy" (giờ trước ngày) khiến ứng dụng không đọc được — bấm để
  tính lại cho các dòng này.
- Nút **Xác định lại Nhóm bệnh (KLN)** dùng cho dữ liệu đã nhập từ trước khi có
  tính năng phân loại Nhóm bệnh (xem mục "Phân loại Nhóm bệnh không lây nhiễm"
  bên dưới) — chỉ gán cho các dòng đang trống, không ghi đè dòng đã có/đã sửa tay.
- Nút **Sao lưu CSDL ngay** / **Mở thư mục sao lưu**: tạo bản sao `benh_nhan.db`
  kèm timestamp trong thư mục `backups/` (tự giữ lại 10 bản gần nhất). Các thao
  tác nguy hiểm (xóa toàn bộ, gộp/xóa bản ghi trùng, sửa lỗi dữ liệu) đều **tự
  động sao lưu trước khi thực hiện** — muốn khôi phục thì đổi tên file backup
  cần dùng thành `benh_nhan.db`.
- Nút **Đặt mật khẩu bảo vệ ứng dụng**: yêu cầu nhập mật khẩu mỗi khi mở app.
  **Lưu ý:** đây chỉ là khóa giao diện, **không mã hóa** file `benh_nhan.db` —
  ai có file đó vẫn mở được bằng công cụ SQLite khác. Phù hợp để tránh người
  không phận sự vô tình mở nhầm, không phù hợp nếu cần bảo mật chống truy cập
  có chủ đích vào file dữ liệu.

### 2. Tab "Danh sách"
- Xem toàn bộ dữ liệu đã nhập, có phân trang (200 dòng/trang).
- Tìm theo họ tên / số CCCD / mã BHYT, lọc theo giới tính, lọc theo Nhóm bệnh
  (KLN) — 1 bệnh nhân mắc nhiều bệnh cùng lúc (vd vừa THA vừa ĐTĐ) vẫn hiện ra
  khi lọc theo từng bệnh riêng lẻ.
- Xuất Excel (.xlsx) hoặc CSV theo đúng bộ lọc đang áp dụng — chọn định dạng
  ngay trong hộp thoại lưu file.

### 3. Tab "Lọc trùng"
- **Tiêu chí xác định trùng**: tích chọn 1 hoặc nhiều trường trong số Số CCCD,
  Mã BHYT, Họ và tên, Năm sinh, Giới tính, Địa chỉ, Phường/Xã, Tỉnh/TP — các
  trường đã chọn được kết hợp bằng AND (tất cả phải khớp mới coi là trùng).
  Số CCCD được chọn sẵn vì chính xác nhất; các tiêu chí khác (nhất là Họ tên +
  Năm sinh) có thể cho trùng giả (2 người khác nhau trùng tên/năm sinh) nên cần
  xem kỹ danh sách chi tiết trước khi gộp.
- Bấm **Quét trùng** để xem danh sách các nhóm bị trùng và số bản ghi dư thừa.
- Chọn 1 nhóm để xem chi tiết từng bản ghi (có thể tích chọn từng dòng).

**Gộp (khuyến nghị — không mất dữ liệu):** khi gộp, ứng dụng giữ lại 1 "bản ghi
chính" (chọn theo "Ngày khám mới nhất" hoặc "Bản ghi đầu tiên"), toàn bộ các
lượt khám của những bản ghi còn lại trong nhóm được dồn vào cột **Lịch sử khám
(đã gộp)** của bản ghi chính, rồi các bản ghi thừa mới bị xóa — không có thông
tin nào bị mất.
- **Gộp các bản ghi đã tích thành 1**: gộp riêng các dòng đã tích trong bảng
  chi tiết (dùng khi chỉ một phần của nhóm là trùng thật).
- **Gộp cả nhóm này thành 1 bản ghi**: gộp toàn bộ nhóm đang chọn.
- **Gộp TẤT CẢ nhóm trùng**: gộp một lượt toàn bộ các nhóm đang hiển thị
  (xử lý được hàng nghìn nhóm trong vài giây).
- **Xóa hẳn các bản ghi đã tích (không gộp)**: xóa thật sự, mất dữ liệu — chỉ
  dùng cho các dòng rác/lỗi thật sự, không phải lượt khám hợp lệ.

**Xác nhận không trùng:** nếu một nhóm thực ra là 2 người khác nhau (trùng tên +
năm sinh chẳng hạn), bấm **Xác nhận: đây KHÔNG phải trùng** để loại nhóm đó
khỏi các lần quét sau (theo đúng tổ hợp tiêu chí đang chọn). Xem/hoàn tác qua
**Quản lý danh sách đã xác nhận...**.

**Xuất dữ liệu:** **Xuất danh sách trùng** (tất cả các dòng thuộc nhóm bị
trùng) và **Xuất danh sách đã lọc trùng - duy nhất** (1 dòng/người, kèm cột
lịch sử khám nếu tích "Kèm cột lịch sử khám khi xuất") — đều xuất được ra
Excel (.xlsx) hoặc CSV, không làm thay đổi CSDL (dùng để xem trước kết quả gộp
trước khi thực sự gộp trong CSDL).

Việc gộp/xóa hàng loạt đều tự động sao lưu CSDL trước khi thực hiện (xem tab
"Nhập dữ liệu" → "Mở thư mục sao lưu" nếu cần khôi phục).

### 4. Tab "Truy vấn SQL"

**Trình tạo câu lệnh SQL (không cần biết cú pháp):** bấm vào ô có dấu tích ở
đầu để mở rộng. Chọn các cột muốn hiển thị, thêm điều kiện lọc (chọn trường —
toán tử như "bằng", "chứa", "lớn hơn", "để trống"... — nhập giá trị), tùy chọn
nhóm theo 1 cột (tự thêm đếm số lượng), sắp xếp và giới hạn số dòng, rồi bấm
**Tạo câu lệnh SQL** — câu lệnh sinh ra sẽ tự động điền vào khung soạn thảo bên
dưới và chạy luôn, có thể chỉnh sửa lại nếu cần trước khi chạy lại.

- Chạy các câu lệnh đơn giản để đếm, lọc, thống kê (chỉ cho phép câu lệnh
  `SELECT`, tên bảng dữ liệu là `patients`).
- Có sẵn danh sách "Câu lệnh nhanh": tổng số bản ghi, số bệnh nhân duy nhất
  theo CCCD, thống kê theo giới tính / tỉnh-thành / phường-xã, top chẩn đoán
  phổ biến, thống kê theo năm sinh, tổ hợp Nhóm bệnh (KLN)...
- Có thể gõ câu lệnh SQL tùy ý, ví dụ:
  ```sql
  SELECT tinh_tp, gioi_tinh, COUNT(*) AS so_luong
  FROM patients
  WHERE chan_doan LIKE '%I10%'
  GROUP BY tinh_tp, gioi_tinh
  ORDER BY so_luong DESC;
  ```
- Xuất kết quả truy vấn ra Excel hoặc CSV bằng nút **Xuất kết quả (Excel/CSV)**.

### 5. Tab "Thống kê"

Vẽ biểu đồ trực quan (dùng PyQt6-Charts) trên dữ liệu hiện có trong CSDL, chọn
qua ô "Loại thống kê":
- **Giới tính** — biểu đồ tròn.
- **Nhóm bệnh (KLN)** — biểu đồ cột, đếm theo TỪNG nhóm bệnh riêng lẻ (Tăng
  huyết áp, Đái tháo đường, COPD/Hen phế quản, Ung thư, Tâm thần, Khác) — 1 bệnh nhân
  mắc nhiều bệnh cùng lúc được tính vào tất cả các nhóm liên quan, không chỉ
  1 nhóm duy nhất. Xem mục "Phân loại Nhóm bệnh không lây nhiễm" bên dưới.
- **Tỉnh/Thành phố**, **Phường/Xã** (top 20), **Chẩn đoán** (top 15) — biểu đồ
  cột ngang, xếp hạng theo số lượng giảm dần.
- **Năm sinh theo thập kỷ** — biểu đồ cột theo từng thập kỷ (1960s, 1970s...).

Biểu đồ tự cập nhật khi mở lại tab này sau khi dữ liệu thay đổi (nhập/gộp/xóa
ở tab khác). Bấm **Làm mới** để vẽ lại ngay lập tức.

> **Không có bản đồ địa lý:** tab này chỉ vẽ biểu đồ cột/tròn theo số liệu,
> KHÔNG vẽ bản đồ tô màu theo ranh giới hành chính thật (kể cả sau khi Hải
> Phòng sáp nhập với Hải Dương từ 1/7/2025, còn 114 xã/phường/đặc khu theo mô
> hình chính quyền 2 cấp) — vì hiện chưa có dữ liệu ranh giới hành chính chính
> thức đáng tin cậy để đưa vào ứng dụng (chỉ có dữ liệu tham khảo chưa xác
> thực từ cộng đồng). Nếu có file GeoJSON/KML ranh giới chính thức (ví dụ từ
> Cổng thông tin điện tử TP Hải Phòng), có thể bổ sung bản đồ thật sau.

### Phân loại Nhóm bệnh không lây nhiễm (KLN)

Ứng dụng ban đầu chỉ quản lý riêng bệnh nhân Tăng huyết áp (THA), nay mở rộng
để quản lý chung nhóm **Bệnh không lây nhiễm (KLN/NCD)**. Cột **Chẩn đoán**
vẫn là nội dung gốc tự do từ file Excel; cột **Nhóm bệnh (KLN)** mới được suy
ra từ đó, dùng để lọc/thống kê theo từng loại bệnh.

- Khi nhập Excel, mỗi dòng được tự động dò từ khóa trong Chẩn đoán để gán vào
  1 hoặc nhiều nhóm: **Tăng huyết áp**, **Đái tháo đường**, **COPD / Hen phế
  quản**, **Ung thư**, **Tâm thần** — hoặc **Khác** nếu Chẩn đoán có nội dung
  nhưng không khớp nhóm nào. 1 bệnh nhân có thể thuộc nhiều nhóm cùng lúc (vd
  "Tăng huyết áp, Đái tháo đường") — thực tế nhiều người mắc đồng thời nhiều
  bệnh không lây nhiễm.
- **Đây chỉ là dò từ khóa đơn giản, KHÔNG phải chẩn đoán y khoa chính thức** —
  cần xem lại và có thể sửa tay trực tiếp cột "Nhóm bệnh (KLN)" ở tab "Danh
  sách" (qua tab "Truy vấn SQL" chạy `UPDATE` hoặc mở rộng tính năng sửa trực
  tiếp trong bảng sau này) nếu cần độ chính xác cao, ví dụ để báo cáo lên
  tuyến trên.
- Dữ liệu đã nhập từ trước khi có tính năng này (cột Nhóm bệnh còn trống)
  dùng nút **Xác định lại Nhóm bệnh (KLN)** ở tab "Nhập dữ liệu" để gán bù,
  không ghi đè các dòng đã có/đã sửa tay.
- Danh mục nhóm bệnh hiện quản lý được định nghĩa trong `core.py` (biến
  `DISEASE_CATEGORIES`) — muốn thêm nhóm bệnh mới (ví dụ Bệnh thận mạn, Tim
  mạch...) thì sửa trực tiếp danh sách này (mã, nhãn hiển thị, danh sách từ
  khóa không dấu) rồi chạy lại "Xác định lại Nhóm bệnh (KLN)".

### 6. Tab "Mạng LAN" (ẩn trên máy có vai trò Máy chủ)

Dùng khi máy này cần dùng chung dữ liệu với 1 **máy chủ** đã có sẵn trong cùng
mạng LAN nội bộ (xem mục "Máy chủ chia sẻ mạng LAN" bên dưới). Tab này chỉ hỗ
trợ đổi qua lại giữa 2 vai trò **Một máy** ↔ **Máy trạm**.

Máy được cài với vai trò **Máy chủ** sẽ **không có tab này** — vì bản thân nó
luôn đọc/ghi trực tiếp CSDL cục bộ của chính mình (không có gì để "chọn kết
nối tới" ở đây), và cũng không đổi được sang vai trò khác ngay trong ứng dụng
(muốn đổi vai trò thì chạy lại trình cài đặt). Tab "Mạng LAN" được thay bằng
tab **"Máy chủ"** (mục 7 bên dưới) để quản lý việc chia sẻ thay vì để trống.

- **Một máy (mặc định)**: hoạt động độc lập như trước đây, không chia sẻ.
- **Máy trạm**: nhập đúng địa chỉ máy chủ (bấm **Kiểm tra kết nối** để thử
  trước), lưu cài đặt và khởi động lại. Sau đó mọi thao tác (nhập Excel, lọc
  trùng, gộp, truy vấn SQL...) đều đọc/ghi trực tiếp vào CSDL trên máy chủ qua
  mạng LAN — không còn dùng file `benh_nhan.db` cục bộ nữa. Nút "Sao lưu CSDL
  ngay" / "Mở thư mục sao lưu" trên máy trạm sẽ tác động tới bản sao lưu trên
  máy chủ (đường dẫn trả về là đường dẫn trên máy chủ).

**Lưu ý bảo mật:** chế độ này hiện **không yêu cầu mật khẩu** — bất kỳ máy nào
truy cập được vào cổng mạng của máy chủ (cùng mạng LAN) đều đọc/ghi được toàn
bộ dữ liệu bệnh nhân, kể cả chạy được câu lệnh SQL tùy ý (tab "Truy vấn SQL" đã
giới hạn phía máy trạm chỉ cho gõ `SELECT`, nhưng máy chủ không tự kiểm tra lại
điều đó). Chỉ dùng tính năng này trong mạng nội bộ đáng tin cậy (không có
Wi-Fi khách lạ dùng chung).

### 7. Tab "Máy chủ" (thay cho tab "Mạng LAN", chỉ hiện với vai trò Máy chủ)

Chỉ xuất hiện trên máy được cài với vai trò **Máy chủ** (xem mục "Máy chủ chia
sẻ mạng LAN" bên dưới), thay thế hẳn cho tab "Mạng LAN" (không có cả 2 tab
cùng lúc) — cho phép quản lý việc chia sẻ dữ liệu ngay trong ứng dụng chính,
không bắt buộc phải mở thêm biểu tượng khay hệ thống
(`QuanLyBenhNhanTHA-Tray.exe`) nữa (dù vẫn dùng song song được, cả 2 đều gọi
chung 1 API quản trị của Windows Service):

- **Trạng thái & địa chỉ**: hiện dịch vụ đang chia sẻ hay đang dừng, địa chỉ
  `http://IP:cổng` để cung cấp cho các máy trạm (nút **Sao chép** / **Mở
  trong trình duyệt**), cùng 2 nút **Bật chia sẻ** / **Dừng chia sẻ** (cần
  thư viện `pywin32` và có thể cần quyền Administrator).
- **Kết nối đang hoạt động**: danh sách các máy trạm đang kết nối (IP, thời
  điểm kết nối, số giây đang rảnh) — chọn 1 dòng rồi bấm **Ngắt kết nối đã
  chọn** để buộc đóng phiên đó ngay lập tức. Tự làm mới mỗi 10 giây.
- **Giới hạn IP được phép kết nối (whitelist)**: giống hệt chức năng trong
  tray — mặc định cho phép mọi IP trong mạng LAN, có thể chuyển sang chỉ cho
  phép các IP trong danh sách rồi bấm **Lưu danh sách IP**.

## Máy chủ chia sẻ mạng LAN

Khi nhiều máy trong cùng một trạm y tế cần dùng chung 1 CSDL, **một máy** (máy
"để bàn", luôn bật) đóng vai trò máy chủ. Vai trò này được chọn **ngay trong
bộ cài đặt duy nhất** của ứng dụng (xem mục "Cài đặt máy chủ" bên dưới) — khi
chọn vai trò **Máy chủ**, ứng dụng chính vẫn dùng được giao diện bình thường
trên máy đó như 1 máy đơn lẻ, đồng thời cài thêm 2 thành phần chạy nền để chia
sẻ CSDL đó cho các máy trạm khác qua mạng LAN:

- **`service.py`** (đóng gói thành `QuanLyBenhNhanTHA-Service.exe`) — Windows
  Service thực sự, làm việc chính (chia sẻ dữ liệu qua mạng), tự khởi động
  cùng Windows và tự khởi động lại nếu bị lỗi, kể cả trước khi có ai đăng nhập
  vào máy. Không có giao diện gì — theo đúng bản chất của Windows Service
  (chạy trong phiên hệ thống riêng, tách biệt khỏi màn hình desktop, nên về
  nguyên tắc không thể tự hiện cửa sổ/icon được).
- **`server_tray.py`** (đóng gói thành `QuanLyBenhNhanTHA-Tray.exe`) — "bảng
  điều khiển" nhỏ, chạy trong phiên đăng nhập của người dùng, hiện 1 icon ở
  khay hệ thống (chấm xanh = đang chia sẻ, đỏ = đang dừng) để xem địa chỉ
  IP:cổng hiện tại, bật/dừng dịch vụ, bật/tắt tự khởi động cùng Windows cho
  chính icon tray này (dịch vụ tự khởi động cùng máy độc lập với tray, xem bên
  dưới), xem **danh sách máy trạm đang kết nối** (kèm nút ngắt 1 kết nối cụ
  thể), và **giới hạn IP được phép kết nối** (whitelist) — xem mục "Quản lý
  kết nối & giới hạn IP" bên dưới. Đóng icon tray (nút "Thoát") **không** làm
  dừng việc chia sẻ — dịch vụ vẫn chạy ngầm bình thường.

### Cài đặt máy chủ

Tải **`QuanLyBenhNhanTHA-Setup-vX.Y.Z.exe`** ở trang Releases (bộ cài đặt duy
nhất, dùng chung cho cả 3 vai trò), chạy file đó (Windows sẽ tự hỏi quyền
Administrator — bắt buộc, vì có thể cần cài Windows Service). Ở bước chọn
**"Vai trò của máy này"**, chọn **Máy chủ**. Trình cài đặt sẽ:
1. Hỏi **cổng chia sẻ** (mặc định `8765`, để trống thì cũng dùng `8765`).
2. Cài đặt bình thường như 1 máy đơn lẻ, đồng thời tự cài và bật Windows
   Service, tự tạo `lan_config.json` với cổng đã chọn.
3. Sau khi cài xong, tự mở icon khay hệ thống (`QuanLyBenhNhanTHA-Tray.exe`) —
   **ghi lại địa chỉ IP:cổng hiển thị ở đó** (menu chuột phải icon, hoặc rê
   chuột vào icon) để cung cấp cho các máy trạm. Bấm menu "Khởi động cùng
   Windows" nếu muốn icon tray tự mở lại mỗi lần có người đăng nhập vào máy
   chủ (dịch vụ chia sẻ vẫn luôn chạy dù có tray hay không).

Gỡ cài đặt qua Start Menu hoặc Settings → Apps như phần mềm bình thường (tự
dừng và gỡ Windows Service nếu có, **không** xóa `benh_nhan.db`/`backups/`).

Muốn đổi cổng chia sẻ sau khi đã cài: dừng dịch vụ, sửa số cổng trong
`lan_config.json` bằng Notepad, rồi bật lại dịch vụ (`services.msc` → restart
dịch vụ `QuanLyBenhNhanTHA_Server`).

Muốn đổi vai trò của máy (ví dụ từ Một máy sang Máy chủ, hoặc ngược lại): chạy
lại trình cài đặt và chọn vai trò mới.

### Quản lý kết nối & giới hạn IP (whitelist)

Có 2 cách tương đương (cùng gọi 1 API quản trị của Windows Service, dùng cách
nào cũng được) — tab **"Máy chủ"** ngay trong ứng dụng chính (xem mục 7 ở
trên), hoặc chuột phải icon tray của máy chủ (`server_tray.py` /
`QuanLyBenhNhanTHA-Tray.exe`):

- **"Kết nối đang hoạt động..."** — hiện danh sách các máy trạm đang kết nối
  (địa chỉ IP, thời điểm kết nối, số giây đang rảnh). Chọn 1 dòng rồi bấm
  **"Ngắt kết nối đã chọn"** để buộc máy trạm đó đóng phiên làm việc ngay lập
  tức (máy trạm sẽ thấy lỗi "Phiên kết nối đã hết hạn" ở lần thao tác tiếp
  theo, cần mở lại kết nối).
- **"Quản lý IP được phép kết nối..."** — mặc định máy chủ chấp nhận **mọi**
  IP trong mạng LAN (giống từ trước tới nay). Chọn "Chỉ cho phép các IP trong
  danh sách", thêm các địa chỉ IP LAN cố định của từng máy trạm (ví dụ
  `192.168.1.11`), rồi bấm **Lưu** — mọi IP không có trong danh sách sẽ bị máy
  chủ từ chối kết nối ngay từ bước đầu tiên. Chính máy chủ (`127.0.0.1`) luôn
  được phép, kể cả khi bật whitelist, để không bao giờ tự khoá mình.

Lưu ý: đây là kiểm soát theo **địa chỉ IP**, không phải xác thực người dùng —
nếu máy trạm trong mạng LAN dùng IP động (DHCP không cố định), whitelist có
thể chặn nhầm sau khi IP đổi; nên đặt IP tĩnh cho các máy trạm trước khi bật
tính năng này. Có thể dùng kèm khoá API (`api_key` trong `lan_config.json`)
để tăng thêm 1 lớp bảo vệ độc lập với IP.

### Đóng gói máy chủ (build từ mã nguồn)

Dùng chung `build.bat` với ứng dụng chính — script này build cả 3 thành phần
(ứng dụng chính, Service, Tray) rồi gom vào 1 thư mục `dist\QuanLyBenhNhanTHA\`
duy nhất:
```
build.bat
```
Muốn build luôn file cài đặt (giống mục "Cài đặt máy chủ" ở trên) thì cần cài
[Inno Setup 6](https://jrsoftware.org/isinfo.php) rồi chạy:
```
"C:\Users\<ten_may>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.3.0 setup.iss
```
Kết quả nằm ở `setup_output\QuanLyBenhNhanTHA-Setup-1.3.0.exe` — chạy file đó
và chọn vai trò **Máy chủ** ở bước cài đặt.

### Cập nhật máy chủ lên bản mới

Dùng chung 1 dòng phiên bản (`VERSION.txt`, tag GitHub dạng `vX.Y.Z`) với ứng
dụng chính — không còn dòng cập nhật riêng cho máy chủ.

Trên máy đã cài với vai trò Máy chủ, chuột phải `update.bat` → **Run as
administrator** (bắt buộc với vai trò Máy chủ, vì cần quyền dừng/bật lại
Windows Service). Script (`update.ps1`) sẽ tự nhận ra máy này đang ở vai trò
Máy chủ (nhờ có sẵn `QuanLyBenhNhanTHA-Service.exe`), tự dừng dịch vụ, thay
file ứng dụng chính lẫn Service/Tray bằng bản mới, rồi bật lại dịch vụ — dữ
liệu `benh_nhan.db`, `lan_config.json`, `acl_config.json`, `backups\` không bị
ảnh hưởng. Repo đang Public nên không cần Personal Access Token (để trống khi
được hỏi) — xem mục D bên dưới nếu sau này repo bị chuyển lại về Private.

> **Lưu ý:** phần Windows Service (`service.py`, `server_tray.py`) chỉ chạy
> được trên Windows và cần thư viện `pywin32`/`pystray` — chưa được kiểm thử
> trên máy Windows thật (được viết theo đúng mẫu chuẩn của pywin32/pystray/Inno
> Setup và đã kiểm tra cú pháp, nhưng nên tự kiểm thử kỹ trên 1 máy Windows
> trước khi dùng cho dữ liệu bệnh nhân thật — đặc biệt là bước cài đặt Windows
> Service tự động qua `setup.iss`).

## Cấu trúc dữ liệu trong CSDL (bảng `patients`)

| Cột | Ý nghĩa |
|---|---|
| tt | Số thứ tự gốc trong file Excel |
| ho_ten | Họ và tên |
| gioi_tinh | Giới tính |
| nam_sinh_raw | Ngày sinh (giữ nguyên định dạng gốc, có thể là năm hoặc ngày/tháng/năm) |
| birth_year | Năm sinh đã tách ra dạng số, dùng để lọc trùng theo Họ tên + Năm sinh |
| ma_bhyt | Mã bảo hiểm y tế |
| so_cccd | Số CCCD/CMND |
| dia_chi | Địa chỉ (số nhà, đường, tổ...) |
| phuong_xa | Phường/Xã |
| tinh_tp | Tỉnh/Thành phố |
| ngay_kham_raw | Ngày khám (giữ nguyên định dạng gốc) |
| ngay_kham_date | Ngày khám chuẩn hóa dạng YYYY-MM-DD để sắp xếp/so sánh |
| chan_doan | Chẩn đoán (nội dung gốc, tự do) |
| benh | Nhóm bệnh không lây nhiễm (KLN) — suy ra từ Chẩn đoán bằng từ khóa, có thể nhiều nhóm cách nhau bởi ", " (vd "Tăng huyết áp, Đái tháo đường") |
| benh_kem_theo | Bệnh kèm theo |
| nguon_file | Tên file Excel đã nhập dòng này |
| imported_at | Thời điểm nhập vào CSDL |
| lich_su_kham | Lịch sử các lượt khám đã bị gộp vào bản ghi này (nếu có) |

Bảng `dedup_exceptions` lưu các nhóm đã được xác nhận "KHÔNG phải trùng" (theo
từng tổ hợp tiêu chí), để tab "Lọc trùng" không hiển thị lại ở các lần quét sau.

## Các file khác được ứng dụng tự tạo ra khi dùng

| File / thư mục | Ý nghĩa |
|---|---|
| `benh_nhan.db` | Dữ liệu chính (SQLite) |
| `backups/` | Các bản sao lưu tự động/thủ công của `benh_nhan.db` (giữ 10 bản gần nhất) |
| `app_password.hash` | Mật khẩu bảo vệ ứng dụng đã băm (nếu đã đặt) — xóa file này để gỡ bỏ mật khẩu nếu quên |
| `update_token.txt` | Personal Access Token GitHub dùng để kiểm tra/tải bản cập nhật (chỉ cần nếu repo bị chuyển về Private) |
| `lan_config.json` | Cấu hình chế độ Mạng LAN của máy này (một máy / máy chủ / máy trạm), xem tab "Mạng LAN" |
| `acl_config.json` | (Chỉ trên máy chủ) Danh sách IP được phép kết nối (whitelist), xem mục "Quản lý kết nối & giới hạn IP" |

Tất cả các file/thư mục trên đều **không** được đưa lên GitHub (đã loại trừ
trong `.gitignore`).

## Lưu ý về chất lượng dữ liệu

- File Excel gốc là danh sách theo **lượt khám**, nên một bệnh nhân có nhiều
  lần khám sẽ xuất hiện nhiều dòng với cùng Số CCCD — đây là điều bình thường,
  không phải lỗi. Dùng tab "Lọc trùng" để quy về danh sách duy nhất theo người
  khi cần.
- Với file mẫu đã kiểm tra: khoảng 2,75% số dòng có cột "Giới tính" và "Ngày
  sinh" bị đảo chỗ, và một phần đáng kể có "Ngày khám" ghi theo kiểu "HH:MM
  dd/mm/yyyy" (giờ trước ngày) thay vì "dd/mm/yyyy HH:MM" — cả hai đều là lỗi
  từ file Excel nguồn. Ứng dụng tự động phát hiện và sửa lại (xem tab "Nhập dữ
  liệu", và mục "Báo cáo chất lượng dữ liệu" sau mỗi lần nhập).
- CSDL được **tự động sao lưu** trước mọi thao tác có thể làm mất dữ liệu (xóa
  toàn bộ, gộp/xóa bản ghi trùng, các nút "Sửa lỗi..."). Xem thư mục `backups/`
  (nút "Mở thư mục sao lưu" trong tab "Nhập dữ liệu") nếu cần khôi phục.

---

## Đóng gói & Triển khai sang máy khác

Repo GitHub (public): **https://github.com/Monsterph6/quanlybenhnhantha**

Có 2 vai trò:
- **Người phát triển (bạn)**: sửa code, build, đẩy bản cập nhật lên GitHub Releases.
- **Máy đích**: chỉ cần giải nén 1 lần, sau đó bấm `update.bat` mỗi khi có bản mới — không cần cài Python hay bất kỳ thứ gì khác.

### A. Quy trình cho người phát triển

**1. Lần đầu — đẩy code lên GitHub:**
```
git init
git add .
git commit -m "Khoi tao du an"
git branch -M main
git remote add origin https://github.com/Monsterph6/quanlybenhnhantha.git
git push -u origin main
```
File `.gitignore` đã loại trừ sẵn dữ liệu bệnh nhân (`*.xlsx`, `*.db`, các file
xuất CSV/Excel) — **kiểm tra lại bằng `git status` trước khi commit**, đừng để
lọt dữ liệu bệnh nhân thật lên GitHub dù là repo private.

**2. Mỗi lần có bản cập nhật muốn đẩy lên:**
```
# 1) Sua code, cap nhat so phien ban trong VERSION.txt (vd: 1.0.1)
git add .
git commit -m "Mo ta thay doi"
git push

# 2) Gan tag dung voi VERSION.txt roi day tag len
git tag v1.0.1
git push origin v1.0.1
```
Khi tag `v*.*.*` được đẩy lên, **GitHub Actions** (`.github/workflows/release.yml`)
tự động: build **cả 3 thành phần** (ứng dụng chính, Service, Tray) bằng
PyInstaller trên máy ảo Windows của GitHub, gom vào 1 thư mục
`dist\QuanLyBenhNhanTHA\` duy nhất, đóng gói thành **2 file** rồi tạo 1
**Release** đính kèm cả hai:
- `QuanLyBenhNhanTHA-Setup-vX.Y.Z.exe` — file **cài đặt** duy nhất (dùng Inno
  Setup) cho cả 3 vai trò (Một máy / Máy trạm / Máy chủ), dùng cho lần đầu
  tiên trên máy mới.
- `QuanLyBenhNhanTHA-vX.Y.Z.zip` — bản **portable** (giải nén là chạy được,
  gồm cả file `.exe` của Service/Tray dùng cho vai trò Máy chủ), dùng làm
  nguồn cho `update.bat` tự tải khi có bản mới.

Theo dõi tiến trình tại tab **Actions** trên trang GitHub của repo.

**3. Build thử trên máy mình (không bắt buộc, chỉ để kiểm tra trước khi tag):**
```
build.bat
```
Kết quả nằm ở `dist\QuanLyBenhNhanTHA\QuanLyBenhNhanTHA.exe` (cùng thư mục có
sẵn `QuanLyBenhNhanTHA-Service.exe` / `-Tray.exe` cho vai trò Máy chủ). Muốn
build luôn file cài đặt thì cần cài
[Inno Setup 6](https://jrsoftware.org/isinfo.php) rồi chạy:
```
"C:\Users\<ten_may>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.3.0 setup.iss
```
Kết quả nằm ở `setup_output\QuanLyBenhNhanTHA-Setup-1.3.0.exe`.

> **Lưu ý kỹ thuật quan trọng:** khi đóng gói bằng PyInstaller, dữ liệu
> `benh_nhan.db` phải nằm **cạnh** file `.exe`, không được nằm trong thư mục
> `_internal` (thư mục này bị xóa và thay bằng bản mới mỗi lần `update.ps1`
> chạy). `core.py` đã xử lý đúng việc này (dùng `sys.executable` khi chạy ở
> dạng đã đóng gói) — nếu sau này sửa lại cách xác định `BASE_DIR`, nhớ giữ
> đúng hành vi này để tránh mất dữ liệu người dùng khi họ bấm cập nhật.

### B. Cài đặt lần đầu trên máy đích (không cần Python)

1. Vào trang Releases của repo, tải file **`QuanLyBenhNhanTHA-Setup-vX.Y.Z.exe`**
   mới nhất.
2. Chạy file đó (Windows sẽ hỏi quyền Administrator — luôn cần, để trình cài
   đặt có thể cài Windows Service nếu chọn vai trò Máy chủ), làm theo hướng
   dẫn. Có 1 bước hỏi **"Vai trò của máy này"**:
   - **Một máy** (mặc định) — dùng độc lập, không chia sẻ qua mạng.
   - **Máy trạm** — nhập địa chỉ IP:cổng của 1 máy chủ đã có sẵn trong mạng
     LAN (xem mục "Máy chủ chia sẻ mạng LAN" ở trên); nếu chưa biết chính xác
     ngay, cứ để trống rồi Next, sau đó nhập/sửa lại trong tab "Mạng LAN" của
     ứng dụng.
   - **Máy chủ** — nhập cổng chia sẻ (mặc định `8765`); máy này vẫn dùng được
     giao diện chính bình thường, đồng thời tự cài và bật thêm 1 Windows
     Service chạy ngầm để chia sẻ dữ liệu cho các máy trạm khác.

   Sau khi cài xong có thể tick "Chạy ngay" để mở ứng dụng luôn.
3. Bộ cài đã kèm sẵn `update.bat`, `update.ps1`, `VERSION.txt` — có shortcut
   "Kiểm tra cập nhật" trong Start Menu. Dữ liệu (`benh_nhan.db`, nếu dùng vai
   trò Một máy hoặc Máy chủ) được tạo ngay trong thư mục cài đặt khi nhập
   Excel lần đầu.

### C. Cập nhật lên bản mới trên máy đích

Mở Start Menu → "Quản lý benh nhan THA" → **Kiểm tra cập nhật** (hoặc bấm đúp
`update.bat` trong thư mục cài đặt — trên máy có vai trò **Máy chủ**, chuột
phải chọn **Run as administrator** vì cần quyền dừng/bật lại Windows Service).
Repo đang **Public** nên không cần Personal Access Token — khi được hỏi, cứ để
trống rồi nhấn Enter (xem mục D bên dưới nếu sau này repo bị chuyển lại về
Private).

`update.bat` sẽ tự so sánh phiên bản, tải bản portable (.zip) mới nếu có, và
thay thế file `.exe` + thư mục `_internal` (và cả file Service/Tray nếu máy
này đang ở vai trò Máy chủ) — **dữ liệu `benh_nhan.db`, `lan_config.json`,
`backups\` không bị ảnh hưởng**. Nếu ứng dụng đang mở, script sẽ nhắc đóng lại
trước khi cập nhật (Windows khóa file .exe/.dll đang chạy).

**Thông báo có bản mới ngay khi mở app:** một khi đã cấu hình `update_token.txt`
(dù chỉ cần chạy `update.bat` 1 lần để nhập token), mỗi lần mở ứng dụng sẽ tự
kiểm tra ngầm (không chặn giao diện, không lỗi nếu mất mạng) — nếu có bản mới
hơn sẽ hiện 1 dải thông báo màu vàng ở đầu cửa sổ, nhắc chạy `update.bat`.

### D. Lấy Personal Access Token (chỉ cần nếu repo đang ở chế độ Private)

Repo `quanlybenhnhantha` hiện đang **Public** — `update.bat` hoạt động bình
thường mà **không cần** token, cứ để trống khi được hỏi và nhấn Enter. Phần
dưới đây chỉ cần làm nếu sau này repo bị chuyển lại về Private:

1. Đăng nhập GitHub → vào **Settings → Developer settings → Personal access
   tokens → Fine-grained tokens → Generate new token**.
2. Đặt tên bất kỳ, **Resource owner**: chọn tài khoản của bạn, **Repository
   access**: chọn "Only select repositories" → chọn repo `quanlybenhnhantha`.
3. Ở mục **Permissions → Repository permissions**, cấp quyền **Contents:
   Read-only** (chỉ cần đọc để tải Release).
4. Bấm **Generate token**, sao chép token (chỉ hiện 1 lần) và dán vào khi
   `update.bat` hỏi.
5. Nếu token hết hạn hoặc nhập sai, xóa file `update_token.txt` rồi chạy lại
   `update.bat` để nhập token mới.
