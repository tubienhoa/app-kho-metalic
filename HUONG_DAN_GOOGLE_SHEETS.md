# Hướng dẫn đưa app kho lên Google Sheets (từng bước)

Mục tiêu: dữ liệu trong app nằm trên **Google Sheets** (cơ sở dữ liệu), và mọi thao tác của nhân viên kho **ghi thẳng về đó**.
Thời gian: khoảng 30–45 phút, làm một lần. Giao diện Google thỉnh thoảng đổi chữ/vị trí nút; nếu không thấy đúng tên, tìm nút gần nghĩa nhất.

**Chú giải:** *CSDL* = cơ sở dữ liệu, nơi lưu dữ liệu · *tài khoản dịch vụ (service account)* = một "tài khoản máy" để app đăng nhập Google thay người · *secrets* = tệp chứa thông tin đăng nhập bí mật của app · *ID file* = chuỗi ký tự nhận diện file Google Sheets.

> Lưu ý an toàn: tệp JSON của tài khoản dịch vụ và tệp `secrets.toml` là **chìa khóa vào dữ liệu**. Không gửi cho ai (kể cả trong chat với tôi), không đưa lên GitHub.

---

## Phần 0. Chuẩn bị trên máy (1 lần)

Mở PowerShell, vào thư mục dự án và cài thư viện cần cho Google:

```powershell
cd D:\CLAUDE_DATA_TONKHO\APP_KHO
```
```powershell
python -m pip install -r requirements.txt
```

---

## Phần A. Tạo 2 file Google Sheets trống

**A0. Tạo thư mục chứa file (khuyến nghị):** vào Google Drive (đúng tài khoản Gmail bạn sẽ dùng cho app) → **Mới → Thư mục mới** → đặt tên `APP_KHO` → mở thư mục đó. Mọi file bên dưới tạo **ngay trong thư mục này**.

Trong thư mục `APP_KHO`: **Mới → Google Trang tính** → làm 2 lần:

| Tên file (đặt đúng tên này) | Chứa gì |
|---|---|
| `APP_KHO_DANH_MUC` | Dữ liệu gốc: mã MISA, hàng kho, hệ số quy đổi… (ít thay đổi) |
| `APP_KHO_PHAT_SINH_2026` | Phiếu nhập xuất, tài khoản, nhật ký… (phát sinh hằng ngày) |

**Lấy ID của từng file:** mở file, nhìn thanh địa chỉ trình duyệt:
`https://docs.google.com/spreadsheets/d/` **`1AbC...xyz`** `/edit`
Phần in đậm giữa `/d/` và `/edit` là **ID**. Ghi lại cả hai ID (một của DANH_MUC, một của PHAT_SINH).

Không cần tạo sheet con nào, lệnh ở Phần D sẽ tự tạo.

---

## Phần B. Tạo tài khoản dịch vụ (để app đọc/ghi 2 file)

1. Mở https://console.cloud.google.com và đăng nhập **cùng Gmail** đã tạo 2 file.
2. Ở thanh trên cùng, bấm chọn dự án → **Dự án mới (New project)** → đặt tên `app-kho` → **Tạo**. Chờ vài giây rồi chọn dự án `app-kho`.
3. Bật 2 dịch vụ: menu ☰ → **APIs & Services (API và dịch vụ) → Library (Thư viện)**:
   - tìm **Google Sheets API** → **Enable (Bật)**;
   - tìm **Google Drive API** → **Enable (Bật)**.
4. Tạo tài khoản dịch vụ: ☰ → **IAM & Admin → Service Accounts (Tài khoản dịch vụ) → Create service account (Tạo tài khoản dịch vụ)**:
   - Tên: `app-kho` → **Create and continue**;
   - bước cấp vai trò: **bỏ qua** (Continue) → **Done**.
5. Tạo khóa: bấm vào tài khoản `app-kho` vừa tạo → thẻ **Keys (Khóa) → Add key → Create new key → JSON → Create**. Máy tải về một tệp `.json`. **Cất kỹ tệp này.**
6. Mở tệp `.json` bằng Notepad, tìm dòng `"client_email": "app-kho@....iam.gserviceaccount.com"` và **chép địa chỉ email đó**.
7. **Chia sẻ cho email đó quyền Người chỉnh sửa (Editor)**:
   - **Cách khuyến nghị:** chia sẻ **thư mục `APP_KHO`** một lần (chuột phải thư mục → **Chia sẻ** → dán email → **Người chỉnh sửa** → bỏ dấu "Thông báo cho người dùng" nếu có → **Chia sẻ/Gửi**). Các file nằm trong thư mục tự được cấp quyền, kể cả file phát sinh năm sau.
   - Nếu file không nằm trong thư mục: chia sẻ **từng file** theo cách tương tự (mở file → nút **Chia sẻ**).

---

## Phần C. Điền tệp `secrets.toml`

1. Vào thư mục `D:\CLAUDE_DATA_TONKHO\APP_KHO\.streamlit\`, **sao chép** `secrets.toml.example` thành `secrets.toml` (cùng thư mục).
2. Mở `secrets.toml` bằng Notepad và điền:
   - `sheet_id_danh_muc` = ID file `APP_KHO_DANH_MUC` (Phần A);
   - `sheet_id_phat_sinh` = ID file `APP_KHO_PHAT_SINH_2026`;
   - các dòng dưới `[luu_tru.credentials]`: chép từ tệp `.json` (`project_id`, `private_key_id`, `private_key`, `client_email`, `client_id`).
3. Chú ý `private_key`: giữ nguyên toàn bộ chuỗi, kể cả các ký hiệu `\n`, và để trong dấu nháy kép `"…"`.
4. Lưu tệp. Tệp này đã được chặn không cho đưa lên GitHub.

---

## Phần D. Chạy 4 lệnh để tạo sheet, kiểm tra, nạp danh mục và tồn đầu

Làm lần lượt, mỗi lệnh xong mới chạy lệnh tiếp. Đứng ở thư mục `APP_KHO` (Phần 0).

**D1. Tạo các sheet trong 2 file** (lệnh tự tạo đúng sheet vào đúng file, kèm dòng tiêu đề):

```powershell
python scripts/cong_cu.py khoi_tao
```

**D2. Kiểm tra kết nối thật** (đủ sheet; ghi 1 dòng thử vào `NHAT_KY` rồi đọc lại):

```powershell
python scripts/cong_cu.py kiem_tra_ket_noi
```

Phải thấy dòng cuối `KẾT NỐI ĐẠT: app đọc và ghi được vào CSDL.` Nếu báo lỗi, xem Phần H. **Chưa đạt thì dừng ở đây, đừng làm tiếp.**

**D3. Nạp danh mục** (1.436 hàng, 1.251 mã MISA, hệ số quy đổi, số MISA tham khảo…). Chỉ nạp vào sheet đang trống nên chạy lại không bị nạp chồng:

```powershell
python scripts/cong_cu.py nap_danh_muc danh_muc/DANH_MUC_NAP_APP_TEST.xlsx
```

**D4. Nạp tồn đầu theo số KHO 23/11/2026** (từng quy cách, đúng đơn vị kho):

```powershell
python scripts/nap_ton_dau_kho.py danh_muc/DANH_MUC_NAP_APP_TEST.xlsx
```

Kết quả mong đợi: 3 dòng "Đã ghi … phiếu" (NM1 sổ CMA, NM1 sổ MTL, NM2 sổ MTL) và cuối cùng `ĐỐI CHIẾU: 1417 hàng-nhà máy, LỆCH 0 → KHỚP 100%`. File đối chiếu nằm ở `danh_muc\KIEM_TRA_TON_DAU_KHO.xlsx`.
Lệnh ghi khá nhiều dòng một lúc nên có thể mất vài phút. Nếu bị ngắt giữa chừng, **chạy lại đúng lệnh đó** (không nạp trùng).

---

## Phần E. Chạy app và kiểm tra dữ liệu đã về CSDL

```powershell
python -m streamlit run app.py
```

Trình duyệt mở http://localhost:8501. Vì đã có `secrets.toml`, app tự dùng Google Sheets (không dùng file thử trên máy).

1. Lần đầu app yêu cầu **tạo tài khoản quản trị** (mã, họ tên, PIN 4–8 số). Tạo xong đăng nhập.
2. Vào tab **🛠 QUẢN TRỊ**: cuối trang có dòng **"💾 Mọi thao tác ghi thẳng vào nơi lưu dữ liệu: Google Sheets – …"** kèm tên 2 file. Đó là bằng chứng app đang nối Google Sheets (nếu thấy "File Excel trên máy" nghĩa là chưa nối, quay lại Phần C).
3. Vào **NGƯỜI DÙNG** tạo vài tài khoản nhân viên kho.
4. Thử một thao tác thật rồi **mở file Google Sheets xem dòng mới**:

| Bạn làm trong app | Mở file | Sheet thấy dòng mới |
|---|---|---|
| Lập phiếu (N1, X1…) | PHÁT SINH | `GIAO_DICH` (mỗi dòng hàng 1 hàng) |
| Hủy phiếu | PHÁT SINH | `HUY_PHIEU` |
| Kiểm kê, khóa sổ | PHÁT SINH | `GIAO_DICH` (K1), `KY_KHOA`, `TON_CHOT` |
| Thêm người dùng, đổi PIN | PHÁT SINH | `NGUOI_DUNG` |
| Mọi thao tác trên | PHÁT SINH | `NHAT_KY` (ai làm gì, lúc nào) |
| Thêm hàng, đổi hệ số quy đổi | DANH_MUC | `HANG_KHO`, `QUY_DOI` |
| Gõ tên đối tác mới khi lập phiếu | DANH_MUC | `DOI_TAC` (tự thêm, kèm số phiếu đã sinh ra nó) |
| Đổi khối lượng riêng | DANH_MUC | `KHOI_LUONG_RIENG` |

Mẹo thử nhanh: tab **📥 NHẬP → N1 · Nhập mua → ➕ THÊM MỚI**, ở ô **Đối tác** gõ một tên bất kỳ rồi nhấn **Enter** (danh sách đối tác ban đầu còn trống, tên bạn gõ sẽ được lưu cho lần sau), chọn một hàng, nhập số lượng, bấm **➕ THÊM DÒNG VÀO PHIẾU**, rồi **💾 LƯU PHIẾU**, rồi mở `APP_KHO_PHAT_SINH_2026` → sheet `GIAO_DICH`: dòng cuối là phiếu bạn vừa lập (cùng số phiếu).

**Quy tắc vàng:** các sheet của file PHÁT SINH **không sửa tay**. Phiếu sai thì **hủy** trong app rồi lập lại. Danh mục cũng sửa qua app (tab Quản trị → Danh mục & quy đổi) để có lịch sử và người duyệt.

---

## Phần F. Chia sẻ cho người khác

- **Không chia sẻ cả thư mục `APP_KHO` cho kế toán** (thư mục chứa cả file phát sinh có tài khoản và nhật ký). Chỉ chia sẻ **từng file** như sau.
- `APP_KHO_DANH_MUC`: chia sẻ riêng file này cho kế toán quyền **Người xem (Viewer)** để đối chiếu danh mục.
- `APP_KHO_PHAT_SINH_2026`: **không chia sẻ cho ai** (chỉ bạn và tài khoản dịch vụ). Kế toán lấy số liệu qua file Excel xuất từ app (tab Quản trị → Báo cáo Excel → **TẠO BÁO CÁO**).
- Nhân viên kho **chỉ dùng app**, không cần truy cập Google Sheets.

---

## Phần G. Đưa app lên mạng cho 8 người dùng (khi đã thử ổn)

Xem README mục 5.5: tạo kho GitHub **Private**, đưa mã lên (đã tự chặn `secrets.toml`, tệp `.json`, `du_lieu_thu/`, `danh_muc/*.xlsx`), tạo app trên share.streamlit.io, rồi dán **toàn bộ nội dung `secrets.toml`** vào ô *Advanced settings → Secrets*.

---

## Phần H. Lỗi thường gặp

| Thông báo / hiện tượng | Nguyên nhân và cách xử lý |
|---|---|
| `SpreadsheetNotFound` hoặc `PERMISSION_DENIED / 403` | Chưa chia sẻ file cho `client_email` (Phần B bước 7), hoặc dán nhầm ID file. Chia sẻ lại cho **cả 2 file** với quyền Người chỉnh sửa |
| `API has not been used … or it is disabled` | Chưa bật Google Sheets API hoặc Google Drive API (Phần B bước 3). Bật rồi chờ 1–2 phút |
| `Invalid JWT / invalid_grant / private_key` | `private_key` trong `secrets.toml` bị mất ký tự `\n` hoặc thiếu dấu nháy. Chép lại nguyên chuỗi từ tệp JSON |
| `429 / Quota exceeded` | Google giới hạn số lần gọi mỗi phút. Chờ 1 phút rồi chạy lại |
| `Sheet … đã có nhưng tiêu đề khác mẫu` | Có người sửa tay dòng tiêu đề (dòng 1). Khôi phục đúng tên cột, hoặc tạo file mới |
| `Thiếu sheet_id_danh_muc / sheet_id_phat_sinh` | Chưa điền đủ 2 ID trong `secrets.toml` |
| App báo "File Excel trên máy" | Chưa có `.streamlit\secrets.toml` đúng chỗ, hoặc còn biến môi trường `APP_KHO_FILE`. Đóng PowerShell, mở lại, chạy lại Phần E |

---

## Phần I. Sau này khi có số kiểm kê chuẩn

1. Tạo file mới `APP_KHO_PHAT_SINH_2026` thứ hai (hoặc đặt tên `…_CHINH_THUC`), chia sẻ cho tài khoản dịch vụ, đổi `sheet_id_phat_sinh` trong `secrets.toml` rồi chạy lại D1–D2. File danh mục **giữ nguyên**.
2. Đăng nhập app (tạo lại tài khoản quản trị và nhân viên), vào **KIỂM KÊ** → tải mẫu → điền số kiểm thực tế → tải lên → **GHI PHIẾU KIỂM KÊ**.
3. Cập nhật danh mục thật (mã đã có đủ, hệ số đã cân) bằng cách báo tôi dựng lại file danh mục chính thức (`DANH_MUC_NAP_APP.xlsx`, không có mã tạm).
4. Đầu mỗi năm: tạo file phát sinh mới, tồn đầu lấy từ `TON_CHOT` cuối năm (khóa sổ trong app).

**Về an toàn dữ liệu:** bản sao lưu hằng tuần trên GitHub có sheet `NGUOI_DUNG` (mã PIN đã băm, không lưu PIN thật). Chỉ để kho sao lưu ở chế độ **Private**.
