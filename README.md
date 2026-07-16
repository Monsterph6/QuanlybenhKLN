# Ứng dụng Lọc trùng Danh sách Bệnh nhân THA

Ứng dụng desktop offline (Python + SQLite + giao diện PyQt6) để nhập danh sách
bệnh nhân từ Excel, lưu trữ, lọc trùng và xuất kết quả ra Excel/CSV.

Có 2 thành phần riêng biệt, đóng gói và cài đặt độc lập với nhau:
- **Ứng dụng chính (`app.py`)** — giao diện PyQt6 dùng hằng ngày, cài trên mọi
  máy (kể cả khi chỉ dùng 1 máy duy nhất, không chia sẻ qua mạng).
- **Máy chủ chia sẻ mạng LAN (`service.py` + `server_tray.py`)** — gói riêng,
  KHÔNG dùng PyQt6 nên nhẹ hơn nhiều, chỉ cài trên 1 máy khi cần nhiều máy
  dùng chung 1 CSDL qua mạng nội bộ. Xem mục "Máy chủ chia sẻ mạng LAN" bên dưới.

Mã nguồn gồm:
- `core.py` — tầng dữ liệu: SQLite, đọc/chuẩn hóa Excel, xuất Excel/CSV (không phụ thuộc giao diện).
- `app.py` — giao diện PyQt6 (nhập `core.py` để xử lý dữ liệu).
- `netserver.py` / `netclient.py` — tầng giao tiếp qua mạng LAN dùng chung bởi
  cả máy chủ và máy trạm, chỉ dùng thư viện chuẩn Python.
- `service.py` / `server_tray.py` — thành phần "máy chủ" (Windows Service +
  tray helper), xem mục riêng bên dưới.

## Yêu cầu

- Python 3.8 trở lên (máy này đang có Python 3.10).
- Thư viện `PyQt6`, `PyQt6-Charts` (vẽ biểu đồ ở tab "Thống kê") và `openpyxl`
  (đã có sẵn trên máy này). Nếu máy khác chưa có, chạy:
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

### 1. Tab "Nhập dữ liệu"
- Chọn file Excel (`.xlsx`) và bấm **Nhập vào cơ sở dữ liệu**.
- Ứng dụng tự dò cột theo tiêu đề "Họ và tên" nên vẫn nhập đúng dù có vài dòng
  tiêu đề gộp (merge) ở đầu file.
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
toan tử như "bằng", "chứa", "lớn hơn", "để trống"... — nhập giá trị), tùy chọn
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
  huyết áp, Đái tháo đường, COPD/Hen phế quản, Ung thư, Khác) — 1 bệnh nhân
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
  quản**, **Ung thư** — hoặc **Khác** nếu Chẩn đoán có nội dung nhưng không
  khớp nhóm nào. 1 bệnh nhân có thể thuộc nhiều nhóm cùng lúc (vd "Tăng huyết
  áp, Đái tháo đường") — thực tế nhiều người mắc đồng thời nhiều bệnh không
  lây nhiễm.
- **Đây chỉ là dò từ khóa đơn giản, KHÔNG phải chẩn đoán y khoa chính thức** —
  cần xem lại và có thể sửa tay trực tiếp cột "Nhóm bệnh (KLN)" ở tab "Danh
  sách" (qua tab "Truy vấn SQL" chạy `UPDATE` hoặc mở rộng tính năng sửa trực
  tiếp trong bảng sau này) nếu cần độ chính xác cao, ví dụ để báo cáo lên
  tuyến trên.
- Dữ liệu đã nhập từ trước khi có tính năng này (cột Nhóm bệnh còn trống)
  dùng nút **Xác định lại Nhóm bệnh (KLN)** ở tab "Nhập dữ liệu" để gán bù,
  không ghi đè các dòng đã có/đã sửa tay.
- Danh mục nhóm bệnh hiện quản lý được định nghĩa trong `core.py` (biến
  `DISEASE_CATEGORIES`) — muốn thêm nhóm bệnh mới (ví dụ Tâm thần, Bệnh thận
  mạn...) thì sửa trực tiếp danh sách này (mã, nhãn hiển thị, danh sách từ
  khóa không dấu) rồi chạy lại "Xác định lại Nhóm bệnh (KLN)".

### 6. Tab "Mạng LAN"

Dùng khi máy này cần dùng chung dữ liệu với 1 **máy chủ** đã được cài riêng
trong cùng mạng LAN nội bộ (xem mục "Máy chủ chia sẻ mạng LAN" bên dưới —
máy chủ là 1 gói cài đặt khác, không nằm trong ứng dụng chính này).

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

## Máy chủ chia sẻ mạng LAN (gói cài đặt riêng)

Khi nhiều máy trong cùng một trạm y tế cần dùng chung 1 CSDL, **một máy** (máy
"để bàn", luôn bật) đóng vai trò máy chủ. Máy chủ là **chương trình độc lập**
với ứng dụng chính — không dùng PyQt6 nên nhẹ hơn nhiều lần, chạy hoàn toàn
ngầm (không cửa sổ), tự khởi động cùng Windows và tự khởi động lại nếu bị lỗi,
kể cả trước khi có ai đăng nhập vào máy — vì nó chạy dưới dạng **Windows
Service** thật sự, không phải chỉ là 1 ứng dụng thường được thêm vào Startup.

Gồm 2 chương trình tách biệt:
- **`service.py`** — Windows Service thực sự, làm việc chính (chia sẻ dữ liệu
  qua mạng). Không có giao diện gì — theo đúng bản chất của Windows Service
  (chạy trong phiên hệ thống riêng, tách biệt khỏi màn hình desktop, nên về
  nguyên tắc không thể tự hiện cửa sổ/icon được).
- **`server_tray.py`** — "bảng điều khiển" nhỏ, chạy trong phiên đăng nhập của
  người dùng, hiện 1 icon ở khay hệ thống (chấm xanh = đang chia sẻ, đỏ = đang
  dừng) để xem địa chỉ IP:cổng hiện tại, bật/dừng dịch vụ, bật/tắt tự khởi
  động cùng Windows cho chính icon tray này (dịch vụ tự khởi động cùng máy độc
  lập với tray, xem bên dưới), xem **danh sách máy trạm đang kết nối** (kèm
  nút ngắt 1 kết nối cụ thể), và **giới hạn IP được phép kết nối** (whitelist)
  — xem mục "Quản lý kết nối & giới hạn IP" bên dưới. Đóng icon tray (nút
  "Thoát") **không** làm dừng việc chia sẻ — dịch vụ vẫn chạy ngầm bình thường.

### Cài đặt máy chủ

**Cách 1 — dùng file cài đặt (khuyến nghị, không cần Python trên máy đích):**

Tải **`QuanLyBenhNhanTHA-Server-Setup-server-vX.Y.Z.exe`** ở trang Releases,
chạy file đó (Windows sẽ tự hỏi quyền Administrator — bắt buộc, vì cài Windows
Service cần quyền này). Trình cài đặt sẽ:
1. Hỏi **cổng chia sẻ** (mặc định `8765`, để trống thì cũng dùng `8765`).
2. Tự cài và bật Windows Service, tự tạo `lan_config.json` với cổng đã chọn.
3. Sau khi cài xong, tự mở icon khay hệ thống (`server_tray.py` /
   `QuanLyBenhNhanTHA-Tray.exe`) — **ghi lại địa chỉ IP:cổng hiển thị ở đó**
   (menu chuột phải icon, hoặc rê chuột vào icon) để cung cấp cho các máy
   trạm. Bấm menu "Khởi động cùng Windows" nếu muốn icon tray tự mở lại mỗi
   lần có người đăng nhập vào máy chủ (dịch vụ chia sẻ vẫn luôn chạy dù có
   tray hay không).

Gỡ cài đặt qua Start Menu hoặc Settings → Apps như phần mềm bình thường (tự
dừng và gỡ Windows Service, **không** xóa `benh_nhan.db`/`backups/`).

**Cách 2 — cài thủ công từ mã nguồn hoặc bản portable (nếu không dùng file cài đặt):**

1. Trên máy sẽ làm máy chủ, cài thư viện riêng cho máy chủ (khác với
   `requirements.txt` của ứng dụng chính):
   ```
   pip install -r requirements-server.txt
   ```
2. Chuột phải vào `install_server.bat` → **Run as administrator**. Script sẽ
   tự tạo `lan_config.json` với cổng mặc định `8765` nếu chưa có, cài và bật
   dịch vụ.
3. (Tuỳ chọn) Chạy `server_tray.py` để xem icon trạng thái ở khay hệ thống.

Muốn đổi cổng (cả 2 cách): dừng dịch vụ, sửa số cổng trong `lan_config.json`
bằng Notepad, bật lại dịch vụ (`services.msc` → restart dịch vụ
`QuanLyBenhNhanTHA_Server`, hoặc `uninstall_server.bat` rồi `install_server.bat`
nếu cài theo Cách 2).

Muốn gỡ theo Cách 2: chuột phải `uninstall_server.bat` → **Run as administrator**.
Dữ liệu `benh_nhan.db` và `backups/` không bị xóa ở cả 2 cách.

### Quản lý kết nối & giới hạn IP (whitelist)

Chuột phải icon tray của máy chủ (`server_tray.py` / `QuanLyBenhNhanTHA-Tray.exe`):

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

Tương tự `build.bat` của ứng dụng chính:
```
build_server.bat
```
Kết quả nằm gọn trong 1 thư mục `dist_server\` (gồm cả Service, Tray,
`install_server.bat`, `uninstall_server.bat`, `update_server.bat`). Muốn build
luôn file cài đặt (giống Cách 1 ở trên) thì cần cài
[Inno Setup 6](https://jrsoftware.org/isinfo.php) rồi chạy:
```
"C:\Users\<ten_may>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DMyAppVersion=0.1.0 setup_server.iss
```
Kết quả nằm ở `setup_output\QuanLyBenhNhanTHA-Server-Setup-0.1.0.exe`. Muốn cài
trực tiếp từ `dist_server\` không qua trình cài đặt thì chạy
`dist_server\install_server.bat` (chuột phải → Run as administrator).

### Cập nhật máy chủ lên bản mới

Máy chủ có **dòng phiên bản riêng** với ứng dụng máy trạm — xem `VERSION_SERVER.txt`
(so với `VERSION.txt` của máy trạm) và tag GitHub dạng `server-vX.Y.Z` (so với
`vX.Y.Z` của máy trạm). Hai dòng cập nhật độc lập, không ảnh hưởng lẫn nhau.

Trên máy chủ đã cài từ bản đóng gói (`dist_server\...`), chuột phải
`update_server.bat` → **Run as administrator**. Script sẽ tự dừng dịch vụ, thay
file Service/Tray bằng bản mới, rồi bật lại dịch vụ — dữ liệu `benh_nhan.db`,
`lan_config.json`, `acl_config.json`, `backups\` không bị ảnh hưởng. Repo đang
Public nên không cần Personal Access Token (để trống khi được hỏi) — xem mục
D bên dưới nếu sau này repo bị chuyển lại về Private.

> **Lưu ý:** phần Windows Service (`service.py`, `server_tray.py`,
> `update_server.ps1`, `setup_server.iss`) chỉ chạy được trên Windows và cần
> thư viện `pywin32`/`pystray` — chưa được kiểm thử trên máy Windows thật
> (được viết theo đúng mẫu chuẩn của pywin32/pystray/Inno Setup và đã kiểm tra
> cú pháp, nhưng nên tự kiểm thử kỹ trên 1 máy Windows trước khi dùng cho dữ
> liệu bệnh nhân thật — đặc biệt là bước cài đặt Windows Service tự động qua
> `setup_server.iss`).

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
tự động: build file .exe bằng PyInstaller trên máy ảo Windows của GitHub, đóng
gói thành **2 file** rồi tạo 1 **Release** đính kèm cả hai:
- `QuanLyBenhNhanTHA-Setup-vX.Y.Z.exe` — file **cài đặt** (dùng Inno Setup),
  dùng cho lần đầu tiên trên máy mới.
- `QuanLyBenhNhanTHA-vX.Y.Z.zip` — bản **portable** (giải nén là chạy được),
  dùng làm nguồn cho `update.bat` tự tải khi có bản mới.

Theo dõi tiến trình tại tab **Actions** trên trang GitHub của repo.

**Riêng cho gói máy chủ** (xem mục "Máy chủ chia sẻ mạng LAN" ở trên): quy
trình tương tự nhưng dùng file version/tag khác, để 2 dòng phiên bản độc lập
với nhau (build/update xong máy trạm không bắt máy chủ phải cập nhật theo và
ngược lại):
```
# Sua code, cap nhat so phien ban trong VERSION_SERVER.txt (vd: 0.1.1)
git add .
git commit -m "Mo ta thay doi"
git push

git tag server-v0.1.1
git push origin server-v0.1.1
```
Tag `server-v*.*.*` kích hoạt workflow riêng
(`.github/workflows/release-server.yml`), build `service.py` + `server_tray.py`,
tạo 1 Release riêng đính kèm **2 file** (giống cấu trúc release máy trạm):
- `QuanLyBenhNhanTHA-Server-Setup-server-vX.Y.Z.exe` — file **cài đặt** (dùng
  Inno Setup, script `setup_server.iss`) — tự hỏi cổng LAN, tự cài/bật Windows
  Service, cần quyền Administrator.
- `QuanLyBenhNhanTHA-Server-server-vX.Y.Z.zip` — bản **portable**, dùng làm
  nguồn cho `update_server.bat` tự tải khi có bản mới, hoặc cài thủ công qua
  `install_server.bat`.

**3. Build thử trên máy mình (không bắt buộc, chỉ để kiểm tra trước khi tag):**
```
build.bat
```
Kết quả nằm ở `dist\QuanLyBenhNhanTHA\QuanLyBenhNhanTHA.exe`. Muốn build luôn
file cài đặt thì cần cài [Inno Setup 6](https://jrsoftware.org/isinfo.php) rồi chạy:
```
"C:\Users\<ten_may>\AppData\Local\Programs\Inno Setup 6\ISCC.exe" /DMyAppVersion=1.0.0 setup.iss
```
Kết quả nằm ở `setup_output\QuanLyBenhNhanTHA-Setup-1.0.0.exe`.

> **Lưu ý kỹ thuật quan trọng:** khi đóng gói bằng PyInstaller, dữ liệu
> `benh_nhan.db` phải nằm **cạnh** file `.exe`, không được nằm trong thư mục
> `_internal` (thư mục này bị xóa và thay bằng bản mới mỗi lần `update.ps1`
> chạy). `core.py` đã xử lý đúng việc này (dùng `sys.executable` khi chạy ở
> dạng đã đóng gói) — nếu sau này sửa lại cách xác định `BASE_DIR`, nhớ giữ
> đúng hành vi này để tránh mất dữ liệu người dùng khi họ bấm cập nhật.

### B. Cài đặt lần đầu trên máy đích (không cần Python)

1. Vào trang Releases của repo, tải file **`QuanLyBenhNhanTHA-Setup-vX.Y.Z.exe`**
   mới nhất.
2. Chạy file đó, làm theo hướng dẫn (mặc định cài vào
   `%LOCALAPPDATA%\Programs\QuanLyBenhNhanTHA`, không cần quyền Admin). Có 1
   bước hỏi **"Địa chỉ máy chủ"** — nếu máy này dùng chung dữ liệu với 1 máy
   chủ đã cài trong mạng LAN (xem mục "Máy chủ chia sẻ mạng LAN" ở trên), nhập
   địa chỉ IP:cổng của máy chủ đó vào đây; nếu dùng 1 máy độc lập thì để trống
   rồi Next. Có thể xem/đổi lại sau trong tab "Mạng LAN" của ứng dụng, không
   bắt buộc phải đúng ngay từ bước cài đặt. Sau khi cài xong có thể tick
   "Chạy ngay" để mở ứng dụng luôn.
3. Bộ cài đã kèm sẵn `update.bat`, `update.ps1`, `VERSION.txt` — có shortcut
   "Kiểm tra cập nhật" trong Start Menu. Dữ liệu (`benh_nhan.db`, nếu dùng 1
   máy độc lập) được tạo ngay trong thư mục cài đặt khi nhập Excel lần đầu.

### C. Cập nhật lên bản mới trên máy đích

Mở Start Menu → "Quản lý benh nhan THA" → **Kiểm tra cập nhật** (hoặc bấm đúp
`update.bat` trong thư mục cài đặt). Repo đang **Public** nên không cần
Personal Access Token — khi được hỏi, cứ để trống rồi nhấn Enter (xem mục D
bên dưới nếu sau này repo bị chuyển lại về Private).

`update.bat` sẽ tự so sánh phiên bản, tải bản portable (.zip) mới nếu có, và
thay thế file `.exe` + thư mục `_internal` — **dữ liệu `benh_nhan.db` không bị
ảnh hưởng** vì nó nằm ngoài `_internal`. Nếu ứng dụng đang mở, script sẽ nhắc
đóng lại trước khi cập nhật (Windows khóa file .exe/.dll đang chạy).

**Thông báo có bản mới ngay khi mở app:** một khi đã cấu hình `update_token.txt`
(dù chỉ cần chạy `update.bat` 1 lần để nhập token), mỗi lần mở ứng dụng sẽ tự
kiểm tra ngầm (không chặn giao diện, không lỗi nếu mất mạng) — nếu có bản mới
hơn sẽ hiện 1 dải thông báo màu vàng ở đầu cửa sổ, nhắc chạy `update.bat`.

`update.bat` chỉ kiểm tra Release của **máy trạm** (tag `vX.Y.Z`) — máy chủ có
cơ chế cập nhật riêng (`update_server.bat`, tag `server-vX.Y.Z`), xem mục
"Cập nhật máy chủ lên bản mới" ở trên.

### D. Lấy Personal Access Token (chỉ cần nếu repo đang ở chế độ Private)

Repo `quanlybenhnhantha` hiện đang **Public** — `update.bat`/`update_server.bat`
hoạt động bình thường mà **không cần** token, cứ để trống khi được hỏi và
nhấn Enter. Phần dưới đây chỉ cần làm nếu sau này repo bị chuyển lại về
Private:

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
