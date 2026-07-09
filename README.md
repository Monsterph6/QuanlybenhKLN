# Ứng dụng Lọc trùng Danh sách Bệnh nhân THA

Ứng dụng desktop offline (Python + SQLite + giao diện PyQt6) để nhập danh sách
bệnh nhân từ Excel, lưu trữ, lọc trùng và xuất kết quả ra Excel/CSV.

Mã nguồn gồm 2 file:
- `core.py` — tầng dữ liệu: SQLite, đọc/chuẩn hóa Excel, xuất Excel/CSV (không phụ thuộc giao diện).
- `app.py` — giao diện PyQt6 (nhập `core.py` để xử lý dữ liệu).

## Yêu cầu

- Python 3.8 trở lên (máy này đang có Python 3.10).
- Thư viện `PyQt6` và `openpyxl` (đã có sẵn trên máy này). Nếu máy khác chưa có, chạy:
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
- Nút **Xóa toàn bộ dữ liệu trong CSDL** dùng khi muốn làm sạch để nhập lại từ đầu.
- Nút **Sửa lỗi đảo cột Giới tính / Ngày sinh** dùng để quét và sửa lại các dòng
  bị lỗi này trong dữ liệu đã nhập từ trước (trước khi có tính năng tự sửa khi nhập).

### 2. Tab "Danh sách"
- Xem toàn bộ dữ liệu đã nhập, có phân trang (200 dòng/trang).
- Tìm theo họ tên / số CCCD / mã BHYT, lọc theo giới tính.
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

Nên sao lưu file `benh_nhan.db` (hoặc xuất Excel/CSV) trước khi dùng chức năng
gộp/xóa hàng loạt, đặc biệt khi lọc theo tiêu chí không phải Số CCCD/Mã BHYT.

### 4. Tab "Truy vấn SQL"
- Chạy các câu lệnh đơn giản để đếm, lọc, thống kê (chỉ cho phép câu lệnh
  `SELECT`, tên bảng dữ liệu là `patients`).
- Có sẵn danh sách "Câu lệnh nhanh": tổng số bản ghi, số bệnh nhân duy nhất
  theo CCCD, thống kê theo giới tính / tỉnh-thành / phường-xã, top chẩn đoán
  phổ biến, thống kê theo năm sinh...
- Có thể gõ câu lệnh SQL tùy ý, ví dụ:
  ```sql
  SELECT tinh_tp, gioi_tinh, COUNT(*) AS so_luong
  FROM patients
  WHERE chan_doan LIKE '%I10%'
  GROUP BY tinh_tp, gioi_tinh
  ORDER BY so_luong DESC;
  ```
- Xuất kết quả truy vấn ra Excel hoặc CSV bằng nút **Xuất kết quả (Excel/CSV)**.

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
| chan_doan | Chẩn đoán |
| benh_kem_theo | Bệnh kèm theo |
| nguon_file | Tên file Excel đã nhập dòng này |
| imported_at | Thời điểm nhập vào CSDL |
| lich_su_kham | Lịch sử các lượt khám đã bị gộp vào bản ghi này (nếu có) |

Bảng `dedup_exceptions` lưu các nhóm đã được xác nhận "KHÔNG phải trùng" (theo
từng tổ hợp tiêu chí), để tab "Lọc trùng" không hiển thị lại ở các lần quét sau.

## Lưu ý về chất lượng dữ liệu

- File Excel gốc là danh sách theo **lượt khám**, nên một bệnh nhân có nhiều
  lần khám sẽ xuất hiện nhiều dòng với cùng Số CCCD — đây là điều bình thường,
  không phải lỗi. Dùng tab "Lọc trùng" để quy về danh sách duy nhất theo người
  khi cần.
- Với file mẫu đã kiểm tra: khoảng 2,75% số dòng (262/9.505 bệnh nhân) có cột
  "Giới tính" và "Ngày sinh" bị đảo chỗ ngay trong file Excel gốc (lỗi nhập liệu
  tại nguồn). Ứng dụng tự động phát hiện và sửa lại (xem tab "Nhập dữ liệu").
- Nên sao lưu file `benh_nhan.db` (hoặc xuất CSV) trước khi dùng chức năng
  xóa dữ liệu hoặc gộp/xóa bản ghi trùng.

---

## Đóng gói & Triển khai sang máy khác

Repo GitHub (private): **https://github.com/Monsterph6/quanlybenhnhantha**

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
   `%LOCALAPPDATA%\Programs\QuanLyBenhNhanTHA`, không cần quyền Admin). Sau khi
   cài xong có thể tick "Chạy ngay" để mở ứng dụng luôn.
3. Bộ cài đã kèm sẵn `update.bat`, `update.ps1`, `VERSION.txt` — có shortcut
   "Kiểm tra cập nhật" trong Start Menu. Dữ liệu (`benh_nhan.db`) được tạo ngay
   trong thư mục cài đặt khi nhập Excel lần đầu.

### C. Cập nhật lên bản mới trên máy đích

Mở Start Menu → "Quản lý benh nhan THA" → **Kiểm tra cập nhật** (hoặc bấm đúp
`update.bat` trong thư mục cài đặt). Lần đầu chạy sẽ hỏi **Personal Access
Token** của GitHub (vì repo đang để **Private**) — xem cách lấy token ở mục D
bên dưới. Token chỉ cần nhập 1 lần, được lưu vào `update_token.txt` (không chia
sẻ file này cho ai).

`update.bat` sẽ tự so sánh phiên bản, tải bản portable (.zip) mới nếu có, và
thay thế file `.exe` + thư mục `_internal` — **dữ liệu `benh_nhan.db` không bị
ảnh hưởng** vì nó nằm ngoài `_internal`. Nếu ứng dụng đang mở, script sẽ nhắc
đóng lại trước khi cập nhật (Windows khóa file .exe/.dll đang chạy).

### D. Lấy Personal Access Token (chỉ cần làm 1 lần/máy, vì repo là Private)

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

> Nếu về sau muốn đơn giản hơn (máy đích không cần token), có thể chuyển repo
> sang **Public** trên GitHub (Settings → Danger Zone → Change visibility) —
> khi đó `update.ps1` vẫn hoạt động bình thường dù không có token hợp lệ, vì
> GitHub cho tải Release công khai không cần xác thực.
