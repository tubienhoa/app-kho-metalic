"""Cấu trúc dữ liệu: tên sheet, cột, loại nghiệp vụ.

Mọi sheet chỉ GHI THÊM dòng (append), không sửa/xóa dòng cũ.
"""

# ---- Các sheet trong file Google Sheets (hoặc file Excel chạy thử) ----
SHEETS = {
    # Người dùng: dòng mới nhất của cùng 'ma' là hiệu lực (đổi PIN/khóa = ghi thêm dòng)
    "NGUOI_DUNG": ["ma", "ten", "nha_may", "quyen", "pin_hash", "hoat_dong", "cap_nhat_luc"],
    # Nhà máy + sổ -> kho MISA (vd NM2+MTL -> 152.SX)
    "KHO_MISA": ["nha_may", "so", "kho_misa"],
    "MA_MISA": ["so", "ma_misa", "ten_misa", "dvt_misa"],
    # Hàng kho = từng quy cách; gắn 1 mã MISA của 1 sổ
    "HANG_KHO": ["ma_hang", "ten_hang", "nhom", "dvt_kho", "so", "ma_misa",
                 "vat_lieu", "quy_cach", "hoat_dong", "ghi_chu", "nguon"],
    # Hệ số: 1 ĐVT kho = he_so ĐVT MISA. Chỉ dùng dòng có người duyệt.
    # Mỗi quy cách có thể có nhiều dòng theo thời gian: dòng hiệu lực mới nhất (<= ngày phiếu) được dùng.
    # hinh_dang/vat_lieu/klr/kt_a..kt_d/kg_ly_thuyet: dữ liệu tính kg theo KÍCH THƯỚC (mm); kg_can: kg cân thực tế 1 ĐVT kho.
    "QUY_DOI": ["ma_hang", "dvt_kho", "dvt_misa", "he_so", "cach_tinh",
                "nguoi_duyet", "ngay_duyet", "hieu_luc_tu", "ghi_chu",
                "hinh_dang", "vat_lieu", "klr", "kt_a", "kt_b", "kt_c", "kt_d", "kg_ly_thuyet", "kg_can"],
    # Khối lượng riêng (g/cm3) theo vật liệu: dòng mới nhất của mỗi vật liệu được dùng. Chỉ quản trị sửa.
    "KHOI_LUONG_RIENG": ["vat_lieu", "klr", "ghi_chu", "nguoi", "thoi_diem"],
    "DONG_NGHIA_DVT": ["dvt_goc", "dvt_chuan", "ghi_chu"],
    # loai: NCC | GIA_CONG | MUA_PHE_LIEU | KHAC
    "DOI_TAC": ["ma_doi_tac", "ten_doi_tac", "loai", "ghi_chu"],
    "BO_PHAN": ["nha_may", "ten_bo_phan"],
    # Hàng được quản trị "mở" cho 1 nhà máy (ngoài các hàng nhà máy đó đã có phát sinh). Chỉ ghi thêm.
    "HANG_NHA_MAY": ["nha_may", "ma_hang", "nguoi", "thoi_diem"],
    # Mỗi dòng = 1 dòng hàng của phiếu, kèm thông tin đầu phiếu (ghi 1 lần cả phiếu)
    "GIAO_DICH": ["phieu_id", "so_phieu", "loai", "ten_loai", "ngay", "nha_may",
                  "nguoi_lap", "thoi_diem_lap", "doi_tac", "bo_phan", "lsx", "po_kh",
                  "so_chung_tu", "kho_den", "phieu_lien_quan", "ly_do", "ghi_chu_phieu",
                  "dong", "dau", "ma_hang", "ten_hang", "so", "ma_misa", "kho_misa",
                  "dvt_kho", "sl_kho", "he_so", "dvt_misa", "sl_misa", "lo", "ghi_chu_dong", "kg_hoa_don"],
    "HUY_PHIEU": ["phieu_id", "so_phieu", "ly_do", "nguoi_huy", "thoi_diem"],
    "KY_KHOA": ["nha_may", "khoa_den_ngay", "nguoi_khoa", "thoi_diem"],
    "TON_CHOT": ["nha_may", "ngay_chot", "ma_hang", "ten_hang", "so", "ma_misa",
                 "dvt_kho", "sl_kho", "dvt_misa", "sl_misa", "thoi_diem"],
    "NHAT_KY": ["thoi_diem", "nguoi", "hanh_dong", "chi_tiet"],
    # Số tồn kế toán MISA tại 1 ngày, CHỈ để đối chiếu trong báo cáo (không tham gia tính tồn của app).
    "TON_MISA_THAM_KHAO": ["ngay", "so", "kho_misa", "ma_misa", "sl_misa", "dvt_misa", "gia_tri"],
}

# ---- Chia file Google Sheets ----
# DANH_MUC: dữ liệu gốc dùng CHUNG hai nhà máy (ít thay đổi, chỉ quản trị ghi qua app, kế toán được xem; gồm cả DOI_TAC).
# PHAT_SINH: dữ liệu RIÊNG từng nhà máy (người dùng, phiếu, hủy, khóa sổ, tồn chốt, nhật ký; mỗi năm một bộ file).
#   Chế độ tách (khuyến nghị): 3 file = DANH_MUC + PHAT_SINH_NM1 + PHAT_SINH_NM2; mỗi dòng ghi vào file của nhà máy đó,
#   khi đọc app gộp hai file lại (nhân viên chỉ thấy nhà máy của mình, quản trị thấy cả hai).
#   Chế độ cũ vẫn dùng được: 2 file (DANH_MUC + 1 file PHAT_SINH chung) hoặc 1 file cho tất cả.
SHEETS_DANH_MUC = ["KHO_MISA", "MA_MISA", "HANG_KHO", "QUY_DOI", "KHOI_LUONG_RIENG", "DONG_NGHIA_DVT", "DOI_TAC",
                   "BO_PHAN", "HANG_NHA_MAY", "TON_MISA_THAM_KHAO"]
SHEETS_PHAT_SINH = ["NGUOI_DUNG", "GIAO_DICH", "HUY_PHIEU", "KY_KHOA", "TON_CHOT", "NHAT_KY"]
assert set(SHEETS_DANH_MUC) | set(SHEETS_PHAT_SINH) == set(SHEETS) and not set(SHEETS_DANH_MUC) & set(SHEETS_PHAT_SINH)

NUMERIC = {"he_so", "dong", "dau", "sl_kho", "sl_misa", "klr", "kt_a", "kt_b", "kt_c", "kt_d",
           "kg_ly_thuyet", "kg_can", "kg_hoa_don", "gia_tri"}

DUNG_SAI = 0.02   # ±2%: vượt thì cảnh báo (kg hóa đơn so với quy đổi; kg cân so với kg lý thuyết)

NHA_MAY = ["NM1", "NM2"]
SO = ["MTL", "CMA"]
QUYEN_KHO, QUYEN_QUAN_TRI = "KHO", "QUAN_TRI"

# ---- Loại nghiệp vụ ----
# dau: +1 nhập, -1 xuất, 0 = phiếu có cả dòng nhập và xuất
# bat_buoc: các ô đầu phiếu bắt buộc; hien: các ô hiện thêm (không bắt buộc)
LOAI = {
    "N1": dict(ten="Nhập mua", nhom="NHAP", dau=+1, doi_tac="NCC",
               bat_buoc=["doi_tac"], hien=["so_chung_tu", "lsx", "po_kh"]),
    "N2": dict(ten="Trả lại từ sản xuất", nhom="NHAP", dau=+1,
               bat_buoc=["bo_phan"], hien=["lsx", "po_kh"]),
    "N3": dict(ten="Nhận lại hàng gia công", nhom="NHAP", dau=+1, doi_tac="GIA_CONG",
               bat_buoc=["doi_tac"], hien=["so_chung_tu", "lsx", "po_kh"]),
    "N4": dict(ten="Nhận hàng chuyển kho", nhom="NHAP", dau=+1,
               bat_buoc=["phieu_lien_quan"], hien=[]),
    "N5": dict(ten="Nhập phần dư / phế liệu thu hồi", nhom="NHAP", dau=+1,
               bat_buoc=["bo_phan"], hien=["lsx", "po_kh"]),
    "X1": dict(ten="Xuất sản xuất", nhom="XUAT", dau=-1,
               bat_buoc=["lsx", "po_kh", "bo_phan"], hien=[]),
    "X2": dict(ten="Xuất tiêu hao", nhom="XUAT", dau=-1,
               bat_buoc=["bo_phan"], hien=["lsx", "po_kh"]),
    "X3": dict(ten="Xuất gia công ngoài", nhom="XUAT", dau=-1, doi_tac="GIA_CONG",
               bat_buoc=["doi_tac"], hien=["so_chung_tu", "lsx", "po_kh"]),
    "X4": dict(ten="Trả nhà cung cấp", nhom="XUAT", dau=-1, doi_tac="NCC",
               bat_buoc=["doi_tac", "ly_do"], hien=["so_chung_tu"]),
    "X5": dict(ten="Chuyển kho sang nhà máy khác", nhom="XUAT", dau=-1,
               bat_buoc=["kho_den"], hien=[]),
    "X6": dict(ten="Xuất hủy (hàng hỏng)", nhom="XUAT", dau=-1,
               bat_buoc=["ly_do"], hien=[]),
    "X7": dict(ten="Bán phế liệu", nhom="XUAT", dau=-1, doi_tac="MUA_PHE_LIEU",
               bat_buoc=["doi_tac"], hien=["so_chung_tu"]),
    "X8": dict(ten="Đổi quy cách (cắt tấm)", nhom="DAC_BIET", dau=0,
               bat_buoc=["bo_phan"], hien=["lsx", "po_kh"]),
    "X9": dict(ten="Chuyển sổ MTL ↔ CMA", nhom="DAC_BIET", dau=0,
               bat_buoc=["ly_do"], hien=[]),
    "K1": dict(ten="Kiểm kê (điều chỉnh)", nhom="QUAN_TRI", dau=0,
               bat_buoc=[], hien=[]),
}

NHAN_O = {  # nhãn hiển thị các ô đầu phiếu
    "doi_tac": "Đối tác", "bo_phan": "Tổ / bộ phận", "lsx": "Lệnh sản xuất",
    "po_kh": "PO khách hàng", "so_chung_tu": "Số chứng từ (phiếu giao / PO mua)",
    "kho_den": "Chuyển đến nhà máy", "ly_do": "Lý do", "phieu_lien_quan": "Phiếu chuyển kho",
}
