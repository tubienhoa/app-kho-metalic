# APP KHO NM1 – NM2

App nhập – xuất cho bộ phận kho, **tự tính tồn**, xuất Excel cho kế toán.
Dữ liệu lưu trên **Google Sheets**. Giao diện chạy trên **Streamlit**. Dùng tạm cho tới khi công ty có ERP.

## 1. Cấu trúc thư mục

| Đường dẫn | Nội dung |
|---|---|
| `app.py` | Giao diện (đăng nhập, trang chủ, lập phiếu, xem tồn, quản trị) |
| `kho/schema.py` | Danh sách sheet, cột, 15 loại nghiệp vụ |
| `kho/logic.py` | Lập/hủy phiếu, tính tồn, kiểm kê, khóa sổ, kiểm tra danh mục |
| `kho/quydoi.py` | Công thức kg theo KÍCH THƯỚC và khối lượng riêng (tấm, ống, hộp, láp, thép V), đổi qua lại kg ↔ tấm/cây |
| `kho/reports.py` | 3 báo cáo Excel: Nhập, Xuất, Nhập – Xuất – Tồn (có kg hóa đơn và % lệch) |
| `kho/store.py` | Đọc/ghi Google Sheets (chạy thật) hoặc file Excel (chạy thử) |
| `scripts/dung_danh_muc.py` | Dựng file danh mục `danh_muc/DANH_MUC_NAP_APP.xlsx` từ V5 + phiếu kiểm v3 (chỉ đọc nguồn) |
| `scripts/cong_cu.py` | Lệnh quản trị: tạo sheet, mẫu danh mục, nạp danh mục, sao lưu CSV |
| `danh_muc/` | File danh mục đã dựng để nạp vào Google Sheets (kèm sheet kiểm tra, danh sách chưa nạp) |
| `scripts/tao_du_lieu_thu.py` | Tạo dữ liệu GIẢ để thử trên máy |
| `scripts/nap_ton_dau_kho.py` | Nạp tồn đầu theo số kho 23/11 (từng quy cách) + file kiểm tra khớp |
| `tests/test_logic.py` | 76 phép kiểm tra tự động (nghiệp vụ, quy đổi, báo cáo) |
| `tests/test_sheets_gia_lap.py` | 17 phép kiểm tra lớp Google Sheets 2 file bằng Google giả lập |
| `HUONG_DAN_GOOGLE_SHEETS.md` | Hướng dẫn từng bước đưa lên Google Sheets |
| `apps_script/SaoLuu.gs` | Sao lưu hằng đêm sang Google Drive |
| `saoluu_github/` | Mẫu kho GitHub riêng để sao lưu CSV hằng tuần |

## 2. Nghiệp vụ

| Mã | Nghiệp vụ | Ô bắt buộc |
|---|---|---|
| N1 | Nhập mua | Đối tác (NCC) |
| N2 | Trả lại từ sản xuất | Tổ |
| N3 | Nhận lại hàng gia công | Đối tác gia công |
| N4 | Nhận hàng chuyển kho | Bấm "Đã nhận đủ" trên phiếu X5 đang chờ |
| N5 | Nhập phần dư / phế liệu thu hồi | Tổ |
| X1 | Xuất sản xuất | **Lệnh SX, PO khách hàng**, Tổ |
| X2 | Xuất tiêu hao | Tổ |
| X3 | Xuất gia công ngoài | Đối tác gia công |
| X4 | Trả nhà cung cấp | NCC, Lý do |
| X5 | Chuyển kho sang nhà máy khác | Nhà máy nhận |
| X6 | Xuất hủy | Lý do |
| X7 | Bán phế liệu | Người mua |
| X8 | Đổi quy cách (cắt tấm) | Tổ; phải có dòng XUẤT và dòng NHẬP |
| X9 | Chuyển sổ MTL ↔ CMA | Lý do; dòng xuất 1 sổ, dòng nhập sổ kia |
| K1 | Kiểm kê | Chỉ quản trị; tải file số kiểm thực tế lên |

LSX, PO khách hàng, Lô chỉ là **cột ghi chú để thống kê**, không tham gia tính tồn.

Giao diện có thanh tab cố định ở đầu trang: **📥 NHẬP · 📤 XUẤT · 🔁 ĐẶC BIỆT · 📋 TỒN & PHIẾU · 🛠 QUẢN TRỊ**. Tab Quản trị chỉ hiện với người có quyền quản trị.
Giao diện gọn, mọi màn hình chính (danh sách phiếu, lập phiếu) **vừa 1 màn hình laptop 15 inch, không phải cuộn**; nút **💾 LƯU PHIẾU** nằm trên đầu trang.
**Bộ lọc phiếu:** Nhà máy · Từ ngày · Đến ngày · Người lập · Tìm. **Mặc định lấy 15 ngày gần nhất** (hôm nay 27 thì từ 12 đến 27). Gõ ngày theo dạng `12/11/2026` (hoặc `12/11`, thiếu năm thì lấy năm nay) rồi **nhấn Enter ở bất kỳ ô nào, hoặc bấm 🔍 LỌC** là lọc lại ngay. Ô ngày dùng ô gõ chữ vì ô chọn ngày của Streamlit không nhận Enter.
Luồng của mỗi nghiệp vụ (N1…X9): bấm thẻ → **danh sách phiếu của loại đó** (bên trái là danh sách bấm-để-xem, bên phải là chi tiết phiếu, có lọc như trên) → bấm **➕ THÊM MỚI** → mới vào màn hình lập phiếu. Lưu xong có 3 nút: **Lập thêm phiếu**, **Xem phiếu vừa lập**, **Về danh sách**. Riêng N4 (nhận hàng chuyển kho), nút ở góc là **➕ NHẬN HÀNG (n chờ)**, và bị khóa khi không có phiếu chuyển kho nào chờ nhận.
Muốn xem phiếu của mọi loại cùng lúc: tab **TỒN & PHIẾU → PHIẾU ĐÃ LẬP**.
Bố cục được căn cho laptop 15 inch: nội dung ở giữa màn hình, rộng tối đa 1240 px.

## 3. Quy tắc chống lỗi dữ liệu

1. **Chỉ ghi thêm dòng.** Không sửa, không xóa dòng cũ. Phiếu sai thì **hủy** (ghi vào sheet `HUY_PHIEU`) rồi lập lại.
2. **Tồn không lưu thành ô.** Tồn = cộng các dòng `GIAO_DICH` chưa hủy (nhập +, xuất −).
3. **Mỗi phiếu ghi 1 lần** và có mã riêng (`phieu_id`). Mạng lỗi thì bấm Lưu lại, phiếu không bị ghi trùng.
4. **Kiểm tra trước khi lưu:**
   - hàng có trong danh mục và đã có hệ số được duyệt;
   - số lượng lớn hơn 0;
   - xuất không vượt tồn;
   - ngày không sau hôm nay và không thuộc kỳ đã khóa;
   - mỗi phiếu chỉ chứa hàng của 1 sổ (trừ X9).
5. **Hệ số quy đổi:** lấy dòng `QUY_DOI` **đã có người duyệt**, có hiệu lực gần nhất. Hệ số được lưu cố định trên từng dòng phiếu, nên đổi hệ số sau này không làm thay đổi phiếu cũ.
6. **Hủy phiếu:**
   - Nhân viên chỉ hủy được phiếu do mình lập trong ngày. Quản trị hủy được mọi phiếu chưa khóa sổ.
   - Không hủy được nếu việc hủy làm tồn bị âm.
7. **Khóa sổ** (chỉ quản trị): sau khi khóa, không ai lập hoặc hủy phiếu có ngày ≤ ngày khóa. Tồn tại ngày khóa được chụp lại vào `TON_CHOT`.
8. **Google Sheets:**
   - Chỉ bạn và tài khoản dịch vụ có quyền sửa file. Nhân viên kho không mở file, chỉ dùng app.
   - **Không sửa tay** các sheet `GIAO_DICH`, `HUY_PHIEU`, `KY_KHOA`, `TON_CHOT`, `NHAT_KY`.
   - Chỉ sửa tay các sheet danh mục: `HANG_KHO`, `MA_MISA`, `QUY_DOI`, `DOI_TAC`, `BO_PHAN`, `KHO_MISA`, `DONG_NGHIA_DVT`.

Giới hạn hiện tại:
- Hai người ở **cùng một kho** bấm xuất **cùng một mã đúng cùng một giây** vẫn có thể làm tồn âm. Trường hợp này rất hiếm; quản trị xem ở "Xem tồn" và lập phiếu điều chỉnh.
- Dữ liệu tạm giữ trong bộ nhớ của app 60 giây, nên màn hình "Xem tồn" có thể chậm tới 1 phút so với nhà máy kia. Lúc lưu phiếu, app luôn đọc dữ liệu mới nhất.

## 4. Chạy thử trên máy (dữ liệu giả)

```powershell
cd D:\CLAUDE_DATA_TONKHO\APP_KHO
```
```powershell
python -m pip install -r requirements.txt
```
```powershell
python scripts/tao_du_lieu_thu.py
```
```powershell
python -m streamlit run app.py
```

Đăng nhập thử: **Quản trị (thử)** PIN `1234` · **Kho NM2 (thử)** PIN `1111` · **Kho NM1 (thử)** PIN `2222`.
Chạy kiểm tra tự động: `python tests/test_logic.py`

## 5. Cài đặt chạy thật (làm 1 lần)

**Hướng dẫn chi tiết từng bước, kèm cách kiểm tra dữ liệu đã về CSDL: xem [HUONG_DAN_GOOGLE_SHEETS.md](HUONG_DAN_GOOGLE_SHEETS.md).** Tóm tắt:

### 5.1 Ba file Google Sheets
Tạo 3 bảng tính trống trong Drive: **`APP_KHO_DANH_MUC`** (dữ liệu gốc dùng chung, gồm cả đối tác), **`APP_KHO_PHAT_SINH_NM1_2026`** và **`APP_KHO_PHAT_SINH_NM2_2026`** (phiếu và hệ thống, mỗi nhà máy một file). Chép ID của từng file (phần giữa `/d/` và `/edit` trong đường link) vào `sheet_id_danh_muc`, `sheet_id_nm1`, `sheet_id_nm2`. Cách chia sheet và lý do: mục 8. (Cách cũ 2 file với `sheet_id_phat_sinh` vẫn chạy được.)

### 5.2 Tài khoản dịch vụ (service account) để app đọc/ghi file
Tạo ở console.cloud.google.com, bật Google Sheets API và Google Drive API, tải khóa JSON, rồi **chia sẻ cả 3 file** cho `client_email` trong file JSON với quyền Người chỉnh sửa. Giữ kín file JSON, không đưa lên GitHub.

### 5.3 Khai thông tin bí mật (secrets)
Chép `.streamlit/secrets.toml.example` thành `.streamlit/secrets.toml`, điền `sheet_id_danh_muc`, `sheet_id_phat_sinh` và các trường từ file JSON. File này đã được chặn trong `.gitignore`.

### 5.4 Tạo sheet, kiểm tra kết nối, nạp danh mục
```powershell
python scripts/cong_cu.py khoi_tao
```
```powershell
python scripts/cong_cu.py kiem_tra_ket_noi
```
```powershell
python scripts/cong_cu.py nap_danh_muc danh_muc/DANH_MUC_NAP_APP_TEST.xlsx
```
`nap_danh_muc` chỉ nạp vào sheet đang trống, không nạp chồng. File `..._TEST.xlsx` có thêm hàng mã tạm cho các dòng chưa có mã; file `DANH_MUC_NAP_APP.xlsx` (không mã tạm) dùng cho bản chính thức. Sau này thêm hàng hoặc đổi hệ số **ngay trong app** (tab Quản trị → Danh mục & quy đổi).

### 5.5 Đưa lên GitHub và Streamlit Cloud
1. Tạo kho GitHub **Private** (ví dụ `APP_KHO`), đưa thư mục này lên. Không đưa `secrets.toml`, file `.json` và `du_lieu_thu/` (đã chặn sẵn trong `.gitignore`).
2. Vào https://share.streamlit.io → **Create app** → chọn kho, file chính là `app.py`.
3. Vào **Advanced settings → Secrets**, dán toàn bộ nội dung `secrets.toml` → **Deploy**.
4. Mở app lần đầu: app yêu cầu tạo tài khoản **quản trị**. Sau đó vào **NGƯỜI DÙNG** để tạo tài khoản cho 8 nhân viên kho.

Lưu ý: app trên Streamlit Cloud có thể "ngủ" khi lâu không ai mở. Lần mở đầu tiên sau đó phải chờ khoảng 1 phút.

### 5.6 Nạp tồn đầu kỳ
- **Dữ liệu thử (hiện tại):** tồn đầu lấy theo số KHO 23/11/2026 từng quy cách:
```powershell
python scripts/nap_ton_dau_kho.py danh_muc/DANH_MUC_NAP_APP_TEST.xlsx
```
  Mỗi nhà máy và sổ một phiếu K1; chạy lại không nạp trùng; ra file `danh_muc/KIEM_TRA_TON_DAU_KHO.xlsx` (tồn kho so với tồn app, chênh phải bằng 0). Dòng kho tồn âm đặt 0 và liệt kê.
- **Bản chính thức (sau này):** đăng nhập quản trị → **KIỂM KÊ** → chọn nhà máy và ngày kiểm → tải mẫu → điền `sl_thuc_te` bằng số kiểm kê thực tế → tải lên → xem chênh lệch → **GHI PHIẾU KIỂM KÊ**. Nên dùng file PHÁT SINH mới (mục 8).

## 6. Sao lưu

1. **Hằng đêm, lên Google Drive.** Làm **cho từng file** (cả DANH_MUC và PHAT_SINH): mở file Sheets → **Tiện ích mở rộng → Apps Script** → dán `apps_script/SaoLuu.gs` → chạy hàm `caiLich` một lần. Từ đó mỗi đêm app tự tạo một bản sao trong thư mục `APP_KHO_SAOLUU`, giữ 30 bản gần nhất.
2. **Hằng tuần, lên GitHub (CSV).** Tạo thêm một kho GitHub **Private** thứ hai, chép vào đó thư mục `saoluu_github/.github`, `kho/`, `scripts/`, rồi khai secrets `GOOGLE_CREDENTIALS`, `SHEET_ID_DANH_MUC`, `SHEET_ID_PHAT_SINH` (xem chú thích trong file `.yml`).
3. **Lịch sử phiên bản của Google Sheets:** vào **Tệp → Nhật ký phiên bản** để xem và khôi phục bản cũ.

Muốn dựng lại dữ liệu: mở bản sao, đổi tên, chia sẻ cho tài khoản dịch vụ, rồi cập nhật mã file tương ứng trong secrets.

## 7. Danh mục và quy đổi kg ↔ tấm/cây

**Nguyên tắc:** kho quản lý từng quy cách (tấm, cây, cái…) dưới một mã kế toán MISA; kế toán ghi theo kg (hoặc đơn vị của MISA). Mỗi quy cách có **hệ số** = số ĐVT MISA trên 1 ĐVT kho (ví dụ 1 tấm thép 3×1220×2500 = 71,8275 kg). Hệ số được lưu trên từng dòng phiếu tại lúc lập, nên đổi hệ số về sau không làm đổi phiếu cũ.

### 7.1 Bảng quy đổi theo kích thước (sheet `QUY_DOI`, `KHOI_LUONG_RIENG`)
- Cột kích thước: `hinh_dang` (tấm / ống / hộp / láp / thép V), `vat_lieu`, `klr` (g/cm³), `kt_a…kt_d` (mm), `kg_ly_thuyet` (app tự tính), `kg_can` (kg cân thực tế nếu có), `he_so` (số đang dùng).
- Thứ tự ưu tiên khi thêm hàng: **cân thực tế đã duyệt > nhập trực tiếp có nguồn > kg lý thuyết theo kích thước**.
- Công thức (mm, g/cm³, kết quả kg/1 đơn vị): tấm `dày×rộng×dài×KLR÷10⁶` · ống `π(D−t)·t·L×KLR÷10⁶` · hộp `2t(a+b−2t)·L×KLR÷10⁶` (chưa tính bo góc) · láp `π/4·d²·L×KLR÷10⁶` · thép V `t(a+b−t)·L×KLR÷10⁶` (chưa tính bo góc).
- Khối lượng riêng mặc định (QT01): thép 7,85 · inox 304/201 7,93 · inox 316 7,98 · inox 430 7,70 · nhôm 2,70 · nhôm A5052 2,68 · đồng 8,93. Sửa trong app, chỉ ảnh hưởng hàng thêm/đổi hệ số sau đó.
- **Chỉ người quản trị** được thêm hàng, đổi hệ số, đổi khối lượng riêng (app chặn cả ở giao diện lẫn trong logic). Mọi người dùng xem được **Máy tính quy đổi** (tab Tồn & Phiếu): nhập số tấm ra kg, hoặc nhập kg ra số tấm kèm phần lẻ.

### 7.2 Dung sai ±2% và kg theo hóa đơn
- Phiếu **nhập mua (N1)** có ô tùy chọn **“Kg ghi trên hóa đơn”**: số kg mà nhà cung cấp ghi trên **hóa đơn (hoặc phiếu cân)** cho dòng hàng đó, để 0 nếu chưa có. Ô **chỉ hiện với hàng kế toán ghi theo kg** (thép tấm, ống, hộp…), không hiện với hàng tính theo cái. App hiện ngay độ lệch so với kg quy đổi; vượt ±2% thì cảnh báo đỏ (không chặn lưu).
- **Đối tác (NCC, nơi gia công, người mua phế liệu)** và **tổ/bộ phận**: một ô duy nhất, chọn trong danh sách **hoặc gõ tên mới rồi nhấn Enter**. Đối tác gõ mới được tự thêm vào sheet `DOI_TAC` (cùng loại) để lần sau chọn lại; gõ lại cùng tên (khác dấu hay hoa/thường) không bị thêm trùng.
- Báo cáo nhập có cột **Kg theo hóa đơn** và **Lệch %**, ô vượt ±2% tô vàng. Tab Quản trị → Kiểm tra danh mục liệt kê các dòng nhập vượt dung sai.
- Khi quản trị nhập **kg cân** lệch quá ±2% so với kg lý thuyết, app cảnh báo; kiểm tra danh mục cũng cảnh báo hệ số lệch kg lý thuyết quá ±2%.

### 7.3 Nạp danh mục từ dữ liệu thật
```powershell
python scripts/dung_danh_muc.py
```
Công cụ đọc `Danh muc du lieu_V5.xlsx` và `So ton kho_23112026_v3.xlsx` (**không sửa**, ghi SHA256 trước/sau) rồi tạo `danh_muc/DANH_MUC_NAP_APP.xlsx` gồm sheet kiểm tra (`TOM_TAT`, `CHUA_NAP` kèm lý do, `CANH_BAO`, `MAP_DONG_V5`) và các sheet danh mục để nạp.
- **Nạp:** dòng kho đã có mã MISA và có hệ số rõ ràng; các dòng trùng quy cách được gộp thành 1 hàng; hàng tồn 0 vẫn nạp.
- **Không nạp** (ghi ở `CHUA_NAP`): 511 dòng chưa có mã (chờ file dữ liệu chuẩn), dòng cần cân thực tế (chờ kiểm kê), dòng khác ĐVT MISA chưa có hệ số, mã chưa có trong báo cáo MISA, ô ghi nhiều mã, hai nguồn trọng lượng lệch >2%.
- Hệ số các dòng nạp được ghi “duyệt theo quy tắc QT01/QT03/QT04/đồng nghĩa 25/11/2026”, nghĩa là duyệt theo quy tắc, chưa phải kế toán/kho xác nhận từng dòng.
- Mã quy cách trong app có dạng `Q00001`; cột `nguon` của `HANG_KHO` ghi dòng nguồn trong V5 để truy vết.
- Danh sách tổ/bộ phận của NM2 (`BO_PHAN`) lấy từ cột “BP/Tổ xuất” của sheet VTTH-NM2, cần kiểm tra lại; NM1 chưa có danh sách (gõ tên trực tiếp vào ô Tổ / bộ phận).

## 8. Bố trí cơ sở dữ liệu trên Google Sheets (3 file: danh mục chung + mỗi nhà máy một file)

| File | Sheet | Ai ghi, ai xem |
|---|---|---|
| **APP_KHO_DANH_MUC** (dữ liệu gốc dùng chung, ít đổi) | `MA_MISA`, `KHO_MISA`, `HANG_KHO`, `QUY_DOI`, `KHOI_LUONG_RIENG`, `DONG_NGHIA_DVT`, `DOI_TAC`, `BO_PHAN`, `HANG_NHA_MAY`, `TON_MISA_THAM_KHAO` | Chỉ quản trị ghi (qua app); riêng `DOI_TAC` app tự thêm khi nhập đối tác mới. Kế toán chỉ xem |
| **APP_KHO_PHAT_SINH_NM1_2026** | `GIAO_DICH`, `HUY_PHIEU`, `KY_KHOA`, `TON_CHOT`, `NHAT_KY`, `NGUOI_DUNG` của NM1 | Chỉ app ghi. Không ai sửa tay. Chỉ chủ file và tài khoản dịch vụ truy cập |
| **APP_KHO_PHAT_SINH_NM2_2026** | các sheet như trên, của NM2 | như trên |

- **Dữ liệu riêng từng nhà máy:** phiếu, hủy phiếu, khóa sổ, tồn chốt, nhật ký và người dùng của nhà máy nào nằm trong file của nhà máy đó. App đọc cả hai file rồi gộp để tính; nhân viên chỉ xem được nhà máy của mình, quản trị xem được cả hai (báo cáo "Cả NM1 + NM2"). Chuyển kho (X5 → N4) ghi ở file bên gửi và file bên nhận. Đối tác dùng chung.
- **Hàng của từng nhà máy:** danh mục hàng vẫn là một bản chung, nhưng mỗi nhà máy chỉ thấy hàng đã có phát sinh hoặc tồn đầu ở nhà máy đó, hoặc hàng quản trị đã "mở" (sheet `HANG_NHA_MAY`; Quản trị → Danh mục hàng & quy đổi → tab Mở hàng cho nhà máy; hàng thêm mới tự mở cho nhà máy được chọn).
- **Chuyển từ 1 file phát sinh chung sang 2 file nhà máy:** `python scripts/chuyen_tach_nha_may.py` (chạy thử trước, thêm `--thuc-hien` để ghi; đọc hướng dẫn đầu file script).

- **Vì sao tách:** làm lại sau khi có số chuẩn chỉ cần tạo file phát sinh mới, danh mục giữ nguyên; kế toán xem được danh mục nhưng không thấy mã PIN (đã băm) và nhật ký; Google giới hạn 10 triệu ô mỗi file và app đọc cả sheet mỗi lần làm mới, nên mỗi năm một file phát sinh (đầu năm sau tồn đầu lấy từ `TON_CHOT` cuối năm).
- `TON_MISA_THAM_KHAO` là số tồn MISA tại 23/11/2026, **chỉ để đối chiếu** trong báo cáo (3 cột cuối của sheet `*_THEO_MA_MISA`), không tham gia tính tồn.
- Tồn đầu của dữ liệu thử lấy theo **số kho từng quy cách**. Khi xuất báo cáo, quy cách được gộp theo mã kế toán (sheet `*_THEO_MA_MISA`, `*_TONG_HOP_THEO_MA`).
- Hàng chưa có mã nằm dưới mã tạm `CHUA_CO_MA` (nhóm “CHƯA CÓ MÃ (TẠM)”), chỉ có trong file danh mục TEST.
