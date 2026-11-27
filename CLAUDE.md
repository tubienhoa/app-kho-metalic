# CLAUDE.md – App nhập xuất tồn kho NM1/NM2

App cho bộ phận kho (NM1, NM2) nhập/xuất hàng, tự tính tồn, xuất Excel cho kế toán. Streamlit + Google Sheets (2 file). Dùng tạm đến khi công ty có ERP.

## Đọc đầu tiên
- **`_ghi_chu_claude/CHECKPOINT_APP_KHO.md`**: trạng thái đầy đủ (quyết định đã chốt, kiến trúc, dữ liệu đã nạp, thông tin Google, việc còn mở, bài học kỹ thuật). Tệp chỉ có trên máy này (đã chặn khỏi git). Đây là BẢN CHÍNH; bản cũ ở `..\_checkpoint\` chỉ là bản sao ngày 27/11/2026.
- `README.md` (cấu trúc, quy tắc, vận hành) và `HUONG_DAN_GOOGLE_SHEETS.md` (hướng dẫn từng bước cho người dùng).
- Dự án đối chiếu danh mục/tồn kho trước đây ở thư mục cha `D:\CLAUDE_DATA_TONKHO\` (file V5, v3, kế toán) là việc RIÊNG. Các file đó chỉ ĐỌC, không sửa.

## Cách làm việc với người dùng
- Trả lời **tiếng Việt, ngắn gọn, chú giải thuật ngữ/viết tắt** khi xuất hiện lần đầu. Người dùng không phải lập trình viên; hay làm "từng bước" và dán ảnh chụp màn hình.
- **Đề xuất/trao đổi TRƯỚC, được duyệt rồi mới tạo hoặc sửa file lớn.** Chưa rõ thì HỎI. Không bịa; nói rõ cái gì chưa kiểm chứng (đặc biệt phần chạy trên Google thật).
- Người dùng chịu trách nhiệm trước ban giám đốc nên cần **chính xác**: đọc lại kết quả trước khi giao, đối chiếu số liệu, không mặc định kế toán hay kho đúng, không ghép dữ liệu chỉ vì tên gần giống.
- Không sửa/đổi tên/xóa file gốc của kế toán và nhà máy (V5, v3, file MISA…); ghi SHA256 trước/sau khi đọc.
- Giao diện: gọn, chữ rõ, **mọi màn chính vừa 1 màn hình laptop 15 inch (khung kiểm 1536×740), không phải cuộn**; nút lưu ở vị trí cố định.

## Lệnh thường dùng (PowerShell, đứng ở thư mục này)
```powershell
python -m streamlit run app.py                       # dùng Google nếu có .streamlit\secrets.toml
$env:APP_KHO_FILE="du_lieu_thu\APP_KHO_DATA_TEST_TON_KHO.xlsx"; python -m streamlit run app.py   # chạy trên Excel thử, KHÔNG đụng Google
python tests/test_logic.py ; python tests/test_sheets_gia_lap.py
python scripts/cong_cu.py kiem_tra_ket_noi
python scripts/dung_danh_muc.py [--ma-tam]           # dựng file danh mục (chỉ đọc nguồn)
```
Luôn đặt `PYTHONIOENCODING=utf-8` khi chạy Python in tiếng Việt trong terminal Windows.

## Cấu trúc
`app.py` (giao diện) · `kho/` (`schema` cột/sheet/loại NV, `logic` nghiệp vụ, `store` Excel/Google, `quydoi` công thức kg, `reports` 3 báo cáo, `auth`) · `scripts/` (công cụ quản trị, nạp dữ liệu) · `tests/` · `danh_muc/` (file nạp, git-ignored) · `du_lieu_thu/` (dữ liệu thử, git-ignored).

## Quy tắc nghiệp vụ không được phá
- Mọi sheet dữ liệu **chỉ ghi thêm dòng**, không sửa/xóa; tồn = tổng giao dịch chưa hủy (không lưu sẵn); sai thì hủy phiếu hoặc lập phiếu điều chỉnh.
- Hệ số quy đổi **lưu cố định trên từng dòng phiếu** tại lúc lập; đổi hệ số không làm đổi phiếu cũ. Dùng dòng QUY_DOI đã có người duyệt, hiệu lực mới nhất ≤ ngày phiếu. Không mặc định hệ số 1 khi ĐVT khác nhau (trừ quy tắc đồng nghĩa đã duyệt).
- **Thêm hàng, đổi hệ số, đổi khối lượng riêng CHỈ quản trị** (chặn cả trong `logic`, không chỉ giao diện). Kiểm kê K1 chỉ quản trị.
- LSX, PO khách hàng, Lô chỉ là cột ghi chú thống kê, không tham gia tính tồn. Mỗi phiếu chỉ một sổ (MTL hoặc CMA), trừ X9. MTL và CMA là HAI bảng mã riêng.
- Dung sai ±2%: kg hóa đơn vs kg quy đổi, kg cân vs kg lý thuyết → **cảnh báo, không chặn**.
- Ngày phiếu không sau hôm nay, không thuộc kỳ đã khóa. Xuất không vượt tồn (K1 là ngoại lệ).

## Bảo mật (rất quan trọng)
- **Không đọc, in, dán hoặc gửi** nội dung `.streamlit/secrets.toml` và mọi tệp trong `_bi_mat/` (khóa tài khoản dịch vụ Google). Không đưa chúng, `*.json`, `du_lieu_thu/`, `danh_muc/*.xlsx`, `_bi_mat/`, `_ghi_chu_claude/` lên GitHub (đã nằm trong `.gitignore`; kiểm lại trước mỗi lần đưa lên).
- Thao tác GHI vào Google Sheets thật (nạp danh mục, nạp tồn đầu, lập phiếu thử…) chỉ khi người dùng đã duyệt. Thử nghiệm tự động phải dùng `APP_KHO_FILE` (file Excel), không dùng Google.
- Thư mục Drive `APP_KHO` ở chế độ **Hạn chế**, chỉ chủ và tài khoản dịch vụ. Nếu người dùng định chia sẻ rộng (đặc biệt link "bất kỳ ai"), phải cảnh báo; kế toán chỉ được xem RIÊNG file danh mục.

## Bẫy kỹ thuật đã gặp
- Streamlit **không nạp lại module `kho/*`** khi sửa → khởi động lại server (`app.py` thì tự nạp).
- `st.date_input` không nhận Enter → lọc ngày dùng form + ô gõ `dd/mm/yyyy`. Ô chọn + gõ tên mới dùng `selectbox(accept_new_options=True)`.
- openpyxl: `ws.max_row` và `ws[...]` tính lại toàn bảng (rất chậm khi ghi nhiều dòng) → tự đếm dòng.
- Hệ số hàng loạt dùng `he_so_bang`; báo cáo chỉ dựng khi bấm nút.
- `AppTest`: đường dẫn script tuyệt đối; ô chọn có `format_func` thì đặt thẳng `session_state`; `session_state` không có `.get`.
- Shell: `sed` làm mất dấu `\` trong đường dẫn Windows; heredoc lồng nháy dễ lỗi → ghi script ra tệp rồi chạy.
- Chi tiết thêm ở mục 9 của checkpoint.
