"""Kiểm thử logic với dữ liệu giả. Chạy: python tests/test_logic.py"""
import io
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pandas as pd  # noqa: E402
from openpyxl import load_workbook  # noqa: E402

from kho import logic as L  # noqa: E402
from kho import reports as R  # noqa: E402
from scripts.tao_du_lieu_thu import tao  # noqa: E402

QT = dict(ma="QT", quyen="QUAN_TRI", nha_may="NM2")
K2 = dict(ma="K2", quyen="KHO", nha_may="NM2")
K1 = dict(ma="K1", quyen="KHO", nha_may="NM1")
N = 0


def ok(cond, msg):
    global N
    N += 1
    assert cond, msg
    print("  OK", msg)


def loi(fn, chua):
    try:
        fn()
    except L.LoiNghiepVu as e:
        ok(chua in str(e), f"báo lỗi đúng: {e}")
        return
    raise AssertionError(f"Không báo lỗi (cần: {chua})")


def main():
    d = tempfile.mkdtemp()
    s = tao(os.path.join(d, "t.xlsx"))
    D = s.read_all
    P, Dg = L.PhieuNhap, L.DongNhap
    hn = str(L.hom_nay())

    # Tồn đầu bằng kiểm kê
    dem = pd.DataFrame({"ma_hang": ["TT3-1220x2500", "OXY"], "sl_thuc_te": [10, 3]})
    ps, bang = L.phieu_kiem_ke(D(), "NM2", "2026-11-25", dem)
    ok(len(ps) == 1 and len(ps[0].dong) == 2, "kiểm kê tạo 1 phiếu K1 2 dòng")
    loi(lambda: L.lap_phieu(s, D(), K2, ps[0]), "quản trị")
    L.lap_phieu(s, D(), QT, ps[0])
    ok(L.ton_mot(D(), "NM2", "TT3-1220x2500") == 10, "tồn đầu 10 tấm")

    # Nhập mua thiếu đối tác
    loi(lambda: L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg("TT3-1220x2500", 5)])), "Đối tác")
    p = P("N1", hn, "NM2", [Dg("TT3-1220x2500", 5, lo="L01")], doi_tac="NCC thép (thử)")
    sp = L.lap_phieu(s, D(), K2, p)
    ok(sp.startswith("NM2-N1-") and sp.endswith("-K2-001"), f"số phiếu {sp}")
    ok(L.lap_phieu(s, D(), K2, p) == sp, "bấm lưu 2 lần không ghi trùng")
    ok(L.ton_mot(D(), "NM2", "TT3-1220x2500") == 15, "tồn 15 sau nhập")
    g = D()["GIAO_DICH"]
    r = g[g["so_phieu"] == sp].iloc[0]
    ok(abs(r["sl_misa"] - 5 * 71.8275) < 1e-6 and r["kho_misa"] == "152.SX", "quy đổi kg + kho MISA lưu trên dòng")

    # Xuất SX bắt buộc LSX/PO, không vượt tồn
    loi(lambda: L.lap_phieu(s, D(), K2, P("X1", hn, "NM2", [Dg("TT3-1220x2500", 2)], bo_phan="Tổ chấn")), "Lệnh sản xuất")
    loi(lambda: L.lap_phieu(s, D(), K2, P("X1", hn, "NM2", [Dg("TT3-1220x2500", 16)], bo_phan="Tổ chấn",
                                          lsx="LSX1", po_kh="PO1")), "chỉ còn 15")
    loi(lambda: L.lap_phieu(s, D(), K2, P("X1", hn, "NM2", [Dg("TT3-1220x2500", 10), Dg("TT3-1220x2500", 6)],
                                          bo_phan="Tổ chấn", lsx="LSX1", po_kh="PO1")), "chỉ còn 15")
    x1 = L.lap_phieu(s, D(), K2, P("X1", hn, "NM2", [Dg("TT3-1220x2500", 4)], bo_phan="Tổ chấn", lsx="LSX1", po_kh="PO1"))
    ok(L.ton_mot(D(), "NM2", "TT3-1220x2500") == 11, "tồn 11 sau xuất SX")

    # Hệ số chưa duyệt, sai nhà máy, ngày tương lai
    loi(lambda: L.lap_phieu(s, D(), QT, P("N5", hn, "NM2", [Dg("CHUA-HS", 1)], bo_phan="Tổ hàn")), "hệ số")
    loi(lambda: L.lap_phieu(s, D(), K2, P("N5", hn, "NM1", [Dg("OXY", 1)], bo_phan="Tổ hàn")), "chỉ được lập phiếu cho NM2")
    loi(lambda: L.lap_phieu(s, D(), K2, P("N5", "2099-01-01", "NM2", [Dg("OXY", 1)], bo_phan="Tổ hàn")), "sau hôm nay")

    # Một phiếu một sổ
    loi(lambda: L.lap_phieu(s, D(), QT, P("N1", hn, "NM1", [Dg("OXY", 1), Dg("CMA-TT3-1220x2500", 1)],
                                          doi_tac="x")), "MỘT sổ")

    # Đổi quy cách: 1 tấm lớn -> 4 tấm nhỏ
    loi(lambda: L.lap_phieu(s, D(), K2, P("X8", hn, "NM2", [Dg("TT3-1220x2500", 1, dau=-1)], bo_phan="Tổ cắt laser")), "ít nhất")
    L.lap_phieu(s, D(), K2, P("X8", hn, "NM2", [Dg("TT3-1220x2500", 1, dau=-1), Dg("TT3-1220x600", 4, dau=1)],
                              bo_phan="Tổ cắt laser"))
    ok(L.ton_mot(D(), "NM2", "TT3-1220x2500") == 10 and L.ton_mot(D(), "NM2", "TT3-1220x600") == 4, "đổi quy cách đúng")

    # Chuyển kho NM2 -> NM1, NM1 nhận
    x5 = P("X5", hn, "NM2", [Dg("TT3-1220x2500", 3)], kho_den="NM1")
    L.lap_phieu(s, D(), K2, x5)
    ok(len(L.x5_cho_nhan(D(), "NM1")["phieu_id"].unique()) == 1, "NM1 có 1 phiếu chờ nhận")
    n4 = L.phieu_nhan_tu_x5(D(), x5.phieu_id, hn)
    loi(lambda: L.lap_phieu(s, D(), K2, n4), "chỉ được lập phiếu cho NM2")
    L.lap_phieu(s, D(), K1, n4)
    ok(L.ton_mot(D(), "NM1", "TT3-1220x2500") == 3 and L.ton_mot(D(), "NM2", "TT3-1220x2500") == 7, "chuyển kho đúng 2 bên")
    ok(L.x5_cho_nhan(D(), "NM1").empty, "hết phiếu chờ nhận")
    n4b = L.phieu_nhan_tu_x5(D(), x5.phieu_id, hn)
    loi(lambda: L.lap_phieu(s, D(), K1, n4b), "không còn chờ nhận")
    loi(lambda: L.huy_phieu(s, D(), K2, x5.phieu_id, "nhầm"), "Hủy phiếu nhận")

    # Chuyển sổ MTL -> CMA ở NM1
    loi(lambda: L.lap_phieu(s, D(), K1, P("X9", hn, "NM1", [Dg("TT3-1220x2500", 1, dau=-1), Dg("TT3-1220x600", 1, dau=1)],
                                          ly_do="mượn")), "sổ còn lại")
    L.lap_phieu(s, D(), K1, P("X9", hn, "NM1", [Dg("TT3-1220x2500", 1, dau=-1), Dg("CMA-TT3-1220x2500", 1, dau=1)], ly_do="mượn"))
    ok(L.ton_mot(D(), "NM1", "CMA-TT3-1220x2500") == 1, "chuyển sổ đúng")

    # Hủy phiếu
    g = D()["GIAO_DICH"]
    id_x1 = g[g["so_phieu"] == x1]["phieu_id"].iloc[0]
    loi(lambda: L.huy_phieu(s, D(), K1, id_x1, "nhầm"), "do mình lập")
    loi(lambda: L.huy_phieu(s, D(), K2, id_x1, " "), "lý do")
    L.huy_phieu(s, D(), K2, id_x1, "nhầm số lượng")
    ok(L.ton_mot(D(), "NM2", "TT3-1220x2500") == 11, "hủy xuất SX trả lại tồn")
    loi(lambda: L.huy_phieu(s, D(), K2, id_x1, "nhầm"), "đã hủy")
    # hủy phiếu nhập mà hàng đã xuất hết -> chặn
    pn = P("N1", hn, "NM2", [Dg("OXY", 2)], doi_tac="NCC")
    L.lap_phieu(s, D(), K2, pn)
    L.lap_phieu(s, D(), K2, P("X2", hn, "NM2", [Dg("OXY", 5)], bo_phan="Tổ hàn"))
    loi(lambda: L.huy_phieu(s, D(), K2, pn.phieu_id, "nhầm"), "âm")

    # Khóa sổ
    loi(lambda: L.khoa_so(s, D(), K2, "NM2", "2026-11-25"), "quản trị")
    L.khoa_so(s, D(), QT, "NM2", "2026-11-25")
    loi(lambda: L.lap_phieu(s, D(), K2, P("N5", "2026-11-25", "NM2", [Dg("OXY", 1)], bo_phan="Tổ hàn")), "đã khóa")
    ok(len(D()["TON_CHOT"]) == 2, "chụp tồn 2 mã khi khóa sổ")

    # Kiểm kê lần 2: đếm thực tế 9 tấm (sổ 11) -> giảm 2
    ps, bang = L.phieu_kiem_ke(D(), "NM2", hn, pd.DataFrame({"ma_hang": ["TT3-1220x2500"], "sl_thuc_te": [9]}))
    ok(ps[0].dong[0].dau == -1 and ps[0].dong[0].sl_kho == 2, "kiểm kê lệch -2")
    L.lap_phieu(s, D(), QT, ps[0])
    ok(L.ton_mot(D(), "NM2", "TT3-1220x2500") == 9, "tồn = số kiểm")

    # Báo cáo
    data = D()
    b = R.bao_cao_nxt(data, "2026-11-26", hn, "NM2")
    wb = load_workbook(io.BytesIO(b))
    ws = wb["MTL_THEO_QUY_CACH"]
    hdr = [c.value for c in ws[5]]
    rows = {r[2]: dict(zip(hdr, r)) for r in ws.iter_rows(min_row=6, values_only=True)}
    r = rows["TT3-1220x2500"]
    ok(r["Tồn đầu kỳ"] == 10 and r["Tồn cuối kỳ"] == 9, f"NXT tồn đầu 10, cuối 9: {r}")
    ws = wb["MTL_THEO_MA_MISA"]
    hdr = [c.value for c in ws[5]]
    rows = {r[1]: dict(zip(hdr, r)) for r in ws.iter_rows(min_row=6, values_only=True)}
    r = rows["THU_THEPTAM3"]
    ok(abs(r["Tồn cuối kỳ"] - (9 * 71.8275 + 4 * 17.2386)) < 1e-6, "NXT theo mã MISA cộng kg các quy cách")
    for fn in (R.bao_cao_nhap, R.bao_cao_xuat):
        wb = load_workbook(io.BytesIO(fn(data, "2026-11-01", hn)))
        ok(wb.sheetnames == ["MTL", "CMA", "MTL_TONG_HOP_THEO_MA", "CMA_TONG_HOP_THEO_MA"], f"{fn.__name__} có sheet MTL, CMA + 2 sheet tổng hợp theo mã")
    wb = load_workbook(io.BytesIO(R.bao_cao_xuat(data, "2026-11-01", hn)))
    ok(not any(r[1] == x1 for r in wb["MTL"].iter_rows(min_row=6, values_only=True)), "phiếu đã hủy không vào báo cáo")
    ok(L.kiem_tra_danh_muc(data)["Mã"].tolist() == ["CHUA-HS"], "kiểm tra danh mục chỉ ra mã chưa duyệt hệ số")
    # ---- Kg theo hóa đơn + cảnh báo ±2% ----
    from kho import quydoi as Q
    loi(lambda: L.lap_phieu(s, D(), K2, P("X2", hn, "NM2", [Dg("OXY", 1, kg_hoa_don=5)], bo_phan="Tổ hàn")), "chỉ ghi ở phiếu nhập mua")
    L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg("TT3-1220x2500", 10, kg_hoa_don=718.275)], doi_tac="NCC"))   # đúng 10 x 71.8275
    L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg("TT3-1220x2500", 10, kg_hoa_don=733.0)], doi_tac="NCC"))     # +2.05%
    L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg("TT3-1220x2500", 10, kg_hoa_don=704.0)], doi_tac="NCC"))     # -1.99%
    cb = L.canh_bao_kg_hoa_don(D())
    ok(len(cb) == 1 and abs(cb.iloc[0]["lech_pct"] - 2.05) < 0.01, f"chỉ dòng lệch +2,05% bị cảnh báo: {cb['lech_pct'].round(2).tolist()}")
    ok(abs(L.lech_kg_hoa_don(718.275, "kg", 733.0) * 100 - 2.0) < 0.1 and L.lech_kg_hoa_don(10, "cái", 5) is None, "hàm lệch kg hóa đơn")
    dd = D()["GIAO_DICH"]
    ok(dd[dd["kg_hoa_don"] > 0].shape[0] == 3, "kg hóa đơn được lưu trên dòng phiếu")
    wb = load_workbook(io.BytesIO(R.bao_cao_nhap(D(), "2026-11-01", hn, "NM2")))
    hdr = [c.value for c in wb["MTL"][5]]
    ok("Kg theo hóa đơn" in hdr and "Lệch % kg hóa đơn so với quy đổi" in hdr, "báo cáo nhập có cột kg hóa đơn và % lệch")
    j = hdr.index("Lệch % kg hóa đơn so với quy đổi")
    vang = [c.fill.fgColor.rgb for c in wb["MTL"].iter_rows(min_row=6, min_col=j + 1, max_col=j + 1) for c in c if c.value not in (None, "")
            and abs(c.value) > 2]
    ok(len(vang) == 1 and vang[0].endswith("FFE699"), "ô lệch vượt ±2% được tô vàng")

    # ---- Công thức kích thước ----
    ok(abs(Q.kg_ly_thuyet("tấm", 7.85, [3, 1220, 2500]) - 71.8275) < 1e-9, "tấm 3x1220x2500 thép = 71,8275 kg")
    ok(abs(Q.kg_ly_thuyet("láp", 7.85, [10, 6000]) - 3.6990) < 1e-3, "láp phi 10 dài 6 m ≈ 3,699 kg (π/4·10²·6000·7,85/10⁶)")
    ok(abs(Q.kg_ly_thuyet("hộp", 7.85, [40, 40, 2, 6000]) - 2 * 2 * (80 - 4) * 6000 * 7.85 / 1e6) < 1e-9, "hộp 40x40x2x6000")
    ok(Q.tach_nguyen_le(150.0, 71.8275)[0] == 2 and abs(Q.tach_nguyen_le(150.0, 71.8275)[1] - 6.345) < 1e-6, "150 kg = 2 tấm + 6,345 kg")
    ok(abs(Q.sl_tu_kg(718.275, 71.8275) - 10) < 1e-9, "718,275 kg = 10 tấm")
    loi_ct = False
    try:
        Q.kg_ly_thuyet("tấm", 7.85, [3, 1220])
    except ValueError:
        loi_ct = True
    ok(loi_ct, "thiếu kích thước -> lỗi")

    # ---- Danh mục & quy đổi: CHỈ quản trị ----
    kw = dict(so="MTL", ma_misa="THU_THEPTAM3", ten_hang="Thép tấm 5 x 1500 x 6000 (thử)", nhom="Thép tấm", dvt_kho="Tấm",
              cach="KICH_THUOC", hinh_dang="tấm", vat_lieu="thép", kt=[5, 1500, 6000])
    loi(lambda: L.them_hang(s, D(), K2, **kw), "Chỉ quản trị")
    loi(lambda: L.dat_klr(s, D(), K2, "thép", 7.9), "Chỉ quản trị")
    loi(lambda: L.doi_he_so(s, D(), K2, "TT3-1220x2500", cach="TRUC_TIEP", hieu_luc_tu=hn, he_so_nhap=70, ghi_chu="cân"), "Chỉ quản trị")
    ma = L.them_hang(s, D(), QT, **kw)
    ok(ma.startswith("Q"), f"thêm hàng mới -> mã {ma}")
    h = L.he_so(D(), ma, hn)
    ok(abs(h["he_so"] - 5 * 1.5 * 6 * 7.85) < 1e-6 and abs(h["kg_ly_thuyet"] - 353.25) < 1e-6 and h["hinh_dang"] == "tấm", "hệ số từ kích thước = 353,25 kg/tấm, lưu đủ kích thước")
    loi(lambda: L.them_hang(s, D(), QT, **kw), "đã có trong danh mục")
    loi(lambda: L.them_hang(s, D(), QT, **{**kw, "ten_hang": "Hàng khác", "ma_misa": "KHONG_CO"}), "không có trong danh sách mã")
    loi(lambda: L.them_hang(s, D(), QT, **{**kw, "ten_hang": "Hàng khác", "cach": "CUNG_DVT"}), "khác ĐVT MISA")
    loi(lambda: L.them_hang(s, D(), QT, **{**kw, "ten_hang": "Hàng khác", "cach": "TRUC_TIEP", "he_so_nhap": 9}), "ghi nguồn")
    t = L.tinh_quy_doi(D(), dvt_kho="tấm", dvt_misa="kg", cach="KICH_THUOC", hinh_dang="tấm", vat_lieu="thép", kt=[5, 1500, 6000], kg_can=350.0)
    ok(t["he_so"] == 350.0 and not t["canh_bao"], "kg cân 350 (lệch -0,9%) không cảnh báo")
    t = L.tinh_quy_doi(D(), dvt_kho="tấm", dvt_misa="kg", cach="KICH_THUOC", hinh_dang="tấm", vat_lieu="thép", kt=[5, 1500, 6000], kg_can=330.0)
    ok(t["he_so"] == 330.0 and len(t["canh_bao"]) == 1, "kg cân 330 (lệch -6,6%) bị cảnh báo vượt ±2%")
    L.dat_klr(s, D(), QT, "thép", 7.9, "thử")
    ok(L.klr_hien_hanh(D())["thép"] == 7.9 and L.klr_hien_hanh(D())["inox 316"] == 7.98, "khối lượng riêng ghi đè theo sheet, loại khác giữ mặc định")
    loi(lambda: L.dat_klr(s, D(), QT, "thép", 70), "0,5 – 25")
    # đổi hệ số có hiệu lực từ hôm nay: phiếu cũ giữ nguyên
    sp_cu = L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg(ma, 1)], doi_tac="NCC"))
    L.doi_he_so(s, D(), QT, ma, cach="TRUC_TIEP", hieu_luc_tu=hn, he_so_nhap=340.0, ghi_chu="cân thực tế 3 tấm 27/11")
    sp_moi = L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg(ma, 1)], doi_tac="NCC"))
    dd = D()["GIAO_DICH"]
    ok(dd[dd["so_phieu"] == sp_cu]["sl_misa"].iloc[0] == 353.25 and dd[dd["so_phieu"] == sp_moi]["sl_misa"].iloc[0] == 340.0,
       "đổi hệ số: phiếu cũ giữ 353,25, phiếu mới dùng 340")
    loi(lambda: L.doi_he_so(s, D(), QT, ma, cach="TRUC_TIEP", hieu_luc_tu="2099-01-01", he_so_nhap=1, ghi_chu="x"), "sau hôm nay")

    # Đối tác gõ tay mới được ghi nhớ vào DOI_TAC; trùng (không dấu, hoa/thường) thì không thêm
    n0 = len(D()["DOI_TAC"])
    L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg("OXY", 1)], doi_tac="Công ty Thiên Long"))
    dt = D()["DOI_TAC"]
    ok(len(dt) == n0 + 1 and dt.iloc[-1]["ten_doi_tac"] == "Công ty Thiên Long" and dt.iloc[-1]["loai"] == "NCC",
       "đối tác mới tự thêm vào DOI_TAC (loại NCC)")
    L.lap_phieu(s, D(), K2, P("N1", hn, "NM2", [Dg("OXY", 1)], doi_tac="  cong ty THIEN LONG "))
    ok(len(D()["DOI_TAC"]) == n0 + 1, "gõ lại cùng tên (khác dấu/hoa thường) -> không thêm trùng")
    L.lap_phieu(s, D(), K2, P("X3", hn, "NM2", [Dg("OXY", 1)], doi_tac="Công ty Thiên Long"))
    dt = D()["DOI_TAC"]
    ok(len(dt) == n0 + 2 and dt.iloc[-1]["loai"] == "GIA_CONG", "cùng tên nhưng khác loại (gia công) -> thêm riêng theo loại")
    L.lap_phieu(s, D(), K2, P("X2", hn, "NM2", [Dg("OXY", 1)], bo_phan="Tổ hàn"))
    ok(len(D()["DOI_TAC"]) == n0 + 2, "phiếu không có đối tác (X2) không thêm gì")

    # Tính hệ số hàng loạt phải giống tính từng hàng
    bang = L.he_so_bang(D(), hn)
    ok(all(abs(L.he_so(D(), m, hn)["he_so"] - bang.loc[m, "he_so"]) < 1e-12 for m in bang.index), "he_so_bang khớp he_so từng hàng")
    ok(bang.loc[ma, "he_so"] == 340.0, "he_so_bang lấy dòng hiệu lực mới nhất (340)")

    # Đọc ngày người dùng gõ
    from datetime import date as _d
    ok(L.doc_ngay("12/11/2026") == _d(2026, 11, 12), "đọc 12/11/2026")
    ok(L.doc_ngay(" 5-1-2026 ") == _d(2026, 1, 5), "đọc 5-1-2026 (có khoảng trắng)")
    ok(L.doc_ngay("12.11.2026") == _d(2026, 11, 12), "đọc 12.11.2026")
    ok(L.doc_ngay("12/11", 2026) == _d(2026, 11, 12), "thiếu năm -> năm mặc định")
    ok(L.doc_ngay("31/02/2026") is None and L.doc_ngay("abc") is None and L.doc_ngay("") is None
       and L.doc_ngay("12/13/2026") is None and L.doc_ngay("12/11/26") is None, "ngày sai / rỗng / năm 2 chữ số -> None")

    # ---- Dữ liệu tách theo nhà máy: hàng của từng nhà máy, mở hàng, người dùng mới nhất ----
    h2 = set(L.hang_cua_nha_may(D(), "NM2")["ma_hang"])
    h1 = set(L.hang_cua_nha_may(D(), "NM1")["ma_hang"])
    ok("TT3-1220x2500" in h2 and "OXY" in h2, "hàng có tồn đầu / phát sinh ở NM2 thì NM2 thấy")
    ok(ma in h2, "hàng quản trị NM2 vừa thêm tự mở cho NM2")
    ok(ma not in h1, "hàng thêm cho NM2 chưa hiện ở NM1")
    gd_ = D()["GIAO_DICH"]
    ok(h1 <= set(gd_[gd_["nha_may"] == "NM1"]["ma_hang"]) | set(D()["HANG_NHA_MAY"].query("nha_may == 'NM1'")["ma_hang"]),
       "NM1 chỉ thấy hàng đã có phát sinh hoặc được mở")
    loi(lambda: L.mo_hang_cho_nha_may(s, D(), K2, ma, "NM1"), "quản trị")
    L.mo_hang_cho_nha_may(s, D(), QT, ma, "NM1")
    ok(ma in set(L.hang_cua_nha_may(D(), "NM1")["ma_hang"]), "quản trị mở hàng cho NM1 -> NM1 thấy")
    loi(lambda: L.mo_hang_cho_nha_may(s, D(), QT, ma, "NM1"), "đã được mở")
    loi(lambda: L.mo_hang_cho_nha_may(s, D(), QT, ma, "NM3"), "Nhà máy không đúng")
    loi(lambda: L.mo_hang_cho_nha_may(s, D(), QT, "KHONG_CO", "NM2"), "Không tìm thấy hàng")
    ma2 = L.them_hang(s, D(), QT, so="MTL", ma_misa=D()["MA_MISA"].iloc[0]["ma_misa"], ten_hang="Hàng thử mở hai nhà máy",
                      nhom="Thử", dvt_kho=D()["MA_MISA"].iloc[0]["dvt_misa"], cach="CUNG_DVT", mo_cho=["NM1", "NM2"])
    ok(ma2 in set(L.hang_cua_nha_may(D(), "NM1")["ma_hang"]) and ma2 in set(L.hang_cua_nha_may(D(), "NM2")["ma_hang"]),
       "them_hang(mo_cho=[NM1, NM2]) -> cả hai nhà máy thấy")
    du_lieu_cu = {k: v for k, v in D().items() if k != "HANG_NHA_MAY"}   # dữ liệu đọc theo bản cũ, chưa có sheet mới
    ok(set(L.hang_cua_nha_may(du_lieu_cu, "NM2")["ma_hang"]) >= {"TT3-1220x2500", "OXY"},
       "thiếu sheet HANG_NHA_MAY (dữ liệu bản cũ) -> không sập, vẫn lọc theo phát sinh")
    # người dùng: dòng mới nhất theo cap_nhat_luc, không theo thứ tự dòng (2 file nhà máy được gộp)
    s.append("NGUOI_DUNG", [dict(ma="DOI", ten="Đổi", nha_may="NM1", quyen="KHO", pin_hash="x", hoat_dong="1", cap_nhat_luc="2026-11-27 10:00:00")])
    s.append("NGUOI_DUNG", [dict(ma="DOI", ten="Đổi", nha_may="NM2", quyen="KHO", pin_hash="x", hoat_dong="1", cap_nhat_luc="2026-11-27 09:00:00")])
    nd = L.nguoi_dung(D()).set_index("ma")
    ok(nd.loc["DOI", "nha_may"] == "NM1", "người dùng: dòng có cap_nhat_luc mới nhất thắng dù nằm trước trong sheet")
    nh_ky = D()["NHAT_KY"]
    ok((nh_ky["hanh_dong"] == "MO_HANG").sum() == 1, "mở hàng được ghi vào nhật ký")
    print(f"\nTẤT CẢ {N} KIỂM TRA ĐẠT")


if __name__ == "__main__":
    main()
