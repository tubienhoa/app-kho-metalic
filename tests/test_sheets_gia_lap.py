"""Thử lớp SheetsStore (2 file) bằng Google GIẢ LẬP trong bộ nhớ – không cần mạng, không cần tài khoản.

Chỉ kiểm logic định tuyến/đọc/ghi của app. KHÔNG thay thế lần thử thật với Google
(chạy: python scripts/cong_cu.py kiem_tra_ket_noi sau khi điền secrets.toml).
Chạy: python tests/test_sheets_gia_lap.py
"""
import os
import sys
import types

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

N = 0


def ok(cond, msg):
    global N
    N += 1
    assert cond, msg
    print("  OK", msg)


class FakeWS:
    def __init__(self, title):
        self.title, self.values = title, []

    def update(self, values, range_name=None, **kw):
        assert range_name == "A1" and kw.get("value_input_option") == "RAW"
        if not self.values:
            self.values = [[]]
        self.values[0] = list(values[0])

    def freeze(self, rows=0):
        pass

    def row_values(self, n):
        return list(self.values[n - 1]) if len(self.values) >= n else []

    def append_rows(self, values, value_input_option=None, insert_data_option=None, table_range=None):
        assert value_input_option == "RAW" and insert_data_option == "INSERT_ROWS"
        self.values.extend([list(r) for r in values])


class FakeSS:
    def __init__(self, title, url):
        self.title, self.url = title, url
        self._ws = {"Sheet1": FakeWS("Sheet1")}   # file Google mới luôn có sẵn 1 sheet trống

    def worksheets(self):
        return list(self._ws.values())

    def add_worksheet(self, title, rows, cols):
        self._ws[title] = FakeWS(title)
        return self._ws[title]

    def worksheet(self, name):
        return self._ws[name]

    def values_batch_get(self, ranges, params=None):
        out = []
        for r in ranges:
            ws = self._ws[r.strip("'")]
            # Google bỏ các ô trống ở cuối dòng và bỏ cả dòng trống ở cuối: mô phỏng đúng như vậy
            vals = []
            for row in ws.values:
                row = list(row)
                while row and row[-1] == "":
                    row.pop()
                vals.append(row)
            out.append({"values": vals})
        return {"valueRanges": out}


class FakeClient:
    def __init__(self, files):
        self.files = files

    def open_by_key(self, key):
        return self.files[key]


def cai_gspread(files):
    m = types.ModuleType("gspread")
    m.service_account_from_dict = lambda info: FakeClient(files)
    sys.modules["gspread"] = m


def main():
    from kho import logic as L
    P, Dg = L.PhieuNhap, L.DongNhap
    from kho.schema import SHEETS, SHEETS_DANH_MUC, SHEETS_PHAT_SINH
    from kho.store import SheetsStore, mo_store

    DM = FakeSS("APP_KHO_DANH_MUC", "https://docs.google.com/spreadsheets/d/DM")
    PS = FakeSS("APP_KHO_PHAT_SINH_2026", "https://docs.google.com/spreadsheets/d/PS")
    cai_gspread({"DM": DM, "PS": PS})
    s = mo_store(dict(kieu="sheets", sheet_id_danh_muc="DM", sheet_id_phat_sinh="PS", credentials={"x": 1}))
    ok(isinstance(s, SheetsStore), "mo_store ra SheetsStore từ cấu hình 2 file")

    # chưa khởi tạo: kiểm tra kết nối phải báo thiếu sheet
    try:
        s.kiem_tra()
        ok(False, "phải báo thiếu sheet")
    except RuntimeError as e:
        ok("THIẾU" in str(e) and "khoi_tao" in str(e), "chưa khởi tạo -> báo thiếu sheet và chỉ lệnh khoi_tao")

    s.khoi_tao()
    ten_dm = {w.title for w in DM.worksheets()} - {"Sheet1"}
    ten_ps = {w.title for w in PS.worksheets()} - {"Sheet1"}
    ok(ten_dm == set(SHEETS_DANH_MUC) and ten_ps == set(SHEETS_PHAT_SINH), "khởi tạo: mỗi file chỉ có đúng sheet của nó")
    ok(DM.worksheet("HANG_KHO").values[0] == SHEETS["HANG_KHO"] and PS.worksheet("GIAO_DICH").values[0] == SHEETS["GIAO_DICH"],
       "dòng tiêu đề đúng mẫu")
    s.khoi_tao()
    ok(len(PS.worksheet("GIAO_DICH").values) == 1, "khởi tạo lần 2 không ghi đè / không thêm dòng")
    dong = s.kiem_tra()
    ok(any("ĐẠT" in d for d in dong) and len(s.mo_ta()) == 2, f"kiểm tra kết nối đạt, mô tả 2 nơi lưu: {[m['ten'] for m in s.mo_ta()]}")
    ok(PS.worksheet("NHAT_KY").values[-1][2] == "KIEM_TRA_KET_NOI", "dòng thử nằm ở file PHÁT SINH (NHAT_KY)")

    # nạp danh mục tối thiểu rồi chạy nghiệp vụ thật qua SheetsStore
    s.append("KHO_MISA", [dict(nha_may="NM2", so="MTL", kho_misa="152.SX")])
    s.append("MA_MISA", [dict(so="MTL", ma_misa="THEPTAM004", ten_misa="Thép tấm 3 ly", dvt_misa="kg")])
    QT = dict(ma="QT", ten="QT", nha_may="NM2", quyen="QUAN_TRI")
    K2 = dict(ma="K2", ten="K2", nha_may="NM2", quyen="KHO")
    data = s.read_all()
    ma = L.them_hang(s, data, QT, so="MTL", ma_misa="THEPTAM004", ten_hang="Thép tấm 3 x 1220 x 2500mm", nhom="Thép tấm",
                     dvt_kho="tấm", cach="KICH_THUOC", hinh_dang="tấm", vat_lieu="thép", kt=[3, 1220, 2500])
    ok(len(DM.worksheet("HANG_KHO").values) == 2 and len(DM.worksheet("QUY_DOI").values) == 2, "thêm hàng -> ghi vào file DANH MỤC")
    ok(len(PS.worksheet("GIAO_DICH").values) == 1, "thêm hàng không ghi gì vào file phát sinh (ngoài NHAT_KY)")
    hn = str(L.hom_nay())
    P, Dg = L.PhieuNhap, L.DongNhap
    sp = L.lap_phieu(s, s.read_all(), K2, P("N1", hn, "NM2", [Dg(ma, 10, kg_hoa_don=733.0)], doi_tac="NCC"))
    ok(len(PS.worksheet("GIAO_DICH").values) == 2 and len(DM.worksheet("HANG_KHO").values) == 2, "lập phiếu -> ghi vào file PHÁT SINH")
    d = s.read_all()
    r = d["GIAO_DICH"].iloc[0]
    ok(r["so_phieu"] == sp and r["sl_kho"] == 10 and abs(r["sl_misa"] - 718.275) < 1e-6 and r["kg_hoa_don"] == 733.0,
       "đọc lại qua SheetsStore ra đúng số (kể cả khi Google bỏ ô trống cuối dòng)")
    ok(L.ton_mot(d, "NM2", ma) == 10, "tồn tính đúng từ dữ liệu đọc về")
    L.huy_phieu(s, s.read_all(), K2, d["GIAO_DICH"].iloc[0]["phieu_id"], "thử hủy")
    ok(len(PS.worksheet("HUY_PHIEU").values) == 2 and L.ton_mot(s.read_all(), "NM2", ma) == 0, "hủy phiếu -> HUY_PHIEU (file phát sinh), tồn về 0")

    # cột mới thêm ở cuối: khoi_tao nối thêm tiêu đề, không mất dữ liệu cũ
    cu = SHEETS["GIAO_DICH"][:-1]
    PS.worksheet("GIAO_DICH").values[0] = list(cu)
    s.khoi_tao()
    ok(PS.worksheet("GIAO_DICH").values[0] == SHEETS["GIAO_DICH"] and len(PS.worksheet("GIAO_DICH").values) == 2,
       "tiêu đề cũ thiếu cột cuối -> được nối thêm, dữ liệu giữ nguyên")
    PS.worksheet("GIAO_DICH").values[0][3] = "SAI_TIEU_DE"
    try:
        s.khoi_tao()
        ok(False, "phải báo tiêu đề sai")
    except ValueError as e:
        ok("tiêu đề khác mẫu" in str(e), "tiêu đề bị sửa tay -> báo lỗi, không ghi đè")

    # chế độ 1 file (tương thích ngược)
    CHUNG = FakeSS("APP_KHO_DATA", "https://docs.google.com/spreadsheets/d/ALL")
    cai_gspread({"ALL": CHUNG})
    s1 = mo_store(dict(kieu="sheets", sheet_id="ALL", credentials={"x": 1}))
    s1.khoi_tao()
    ok({w.title for w in CHUNG.worksheets()} - {"Sheet1"} == set(SHEETS) and len(s1.mo_ta()) == 1, "chỉ có sheet_id -> 1 file chứa đủ mọi sheet")
    try:
        mo_store(dict(kieu="sheets", sheet_id_danh_muc="DM", credentials={"x": 1}))
        ok(False, "phải báo thiếu cấu hình")
    except ValueError as e:
        ok("sheet_id_phat_sinh" in str(e), "thiếu 1 trong 2 mã file -> báo lỗi rõ ràng")

    # ================= 3 file: DANH MỤC chung + PHÁT SINH riêng từng nhà máy =================
    from kho.schema import NHA_MAY
    DM3 = FakeSS("APP_KHO_DANH_MUC", "https://docs.google.com/spreadsheets/d/DM3")
    F1 = FakeSS("APP_KHO_PHAT_SINH_NM1_2026", "https://docs.google.com/spreadsheets/d/N1")
    F2 = FakeSS("APP_KHO_PHAT_SINH_NM2_2026", "https://docs.google.com/spreadsheets/d/N2")
    cai_gspread({"DM3": DM3, "N1": F1, "N2": F2})
    t = mo_store(dict(kieu="sheets", sheet_id_danh_muc="DM3", sheet_id_nm1="N1", sheet_id_nm2="N2", credentials={"x": 1}))
    ok(t.tach_nha_may and len(t.mo_ta()) == 3, "mo_store: có sheet_id_nm1 + sheet_id_nm2 -> chế độ 3 file tách nhà máy")
    for cfg, chua in [(dict(sheet_id_danh_muc="DM3", sheet_id_nm1="N1"), "sheet_id_nm2"),
                      (dict(sheet_id_danh_muc="DM3", sheet_id_nm1="N1", sheet_id_nm2="N1"), "3 file khác nhau"),
                      (dict(sheet_id_nm1="N1", sheet_id_nm2="N2"), "sheet_id_danh_muc")]:
        try:
            mo_store(dict(kieu="sheets", credentials={"x": 1}, **cfg))
            ok(False, "phải báo lỗi cấu hình")
        except ValueError as e:
            ok(chua in str(e), f"cấu hình sai -> báo lỗi rõ ({chua})")
    try:
        t.kiem_tra()
        ok(False, "phải báo thiếu sheet")
    except RuntimeError as e:
        ok("THIẾU" in str(e), "3 file chưa khởi tạo -> báo thiếu sheet")
    t.khoi_tao()
    td = lambda f: {w.title for w in f.worksheets()} - {"Sheet1"}
    ok(td(DM3) == set(SHEETS_DANH_MUC) and td(F1) == set(SHEETS_PHAT_SINH) and td(F2) == set(SHEETS_PHAT_SINH),
       "khởi tạo: danh mục chỉ có sheet chung (kể cả DOI_TAC, HANG_NHA_MAY); mỗi file nhà máy có đủ sheet phát sinh")
    dong = t.kiem_tra()
    ok(any("ĐẠT" in d for d in dong) and F1.worksheet("NHAT_KY").values[-1][2] == "KIEM_TRA_KET_NOI"
       and F2.worksheet("NHAT_KY").values[-1][2] == "KIEM_TRA_KET_NOI", "kiểm tra kết nối: ghi và đọc lại dòng thử ở CẢ HAI file nhà máy")

    t.append("KHO_MISA", [dict(nha_may="NM2", so="MTL", kho_misa="152.SX"), dict(nha_may="NM1", so="MTL", kho_misa="152.GC")])
    t.append("MA_MISA", [dict(so="MTL", ma_misa="THEPTAM004", ten_misa="Thép tấm 3 ly", dvt_misa="kg")])
    QT1 = dict(ma="QT", ten="QT", nha_may="NM2", quyen="QUAN_TRI")
    K_1 = dict(ma="K1", ten="K1", nha_may="NM1", quyen="KHO")
    K_2 = dict(ma="K2", ten="K2", nha_may="NM2", quyen="KHO")
    mh = L.them_hang(t, t.read_all(), QT1, so="MTL", ma_misa="THEPTAM004", ten_hang="Thép tấm 3 x 1220 x 2500mm", nhom="Thép tấm",
                     dvt_kho="tấm", cach="KICH_THUOC", hinh_dang="tấm", vat_lieu="thép", kt=[3, 1220, 2500], mo_cho=["NM1", "NM2"])
    ok(len(DM3.worksheet("HANG_KHO").values) == 2 and len(DM3.worksheet("HANG_NHA_MAY").values) == 3,
       "thêm hàng -> danh mục chung; mở cho 2 nhà máy -> 2 dòng HANG_NHA_MAY")
    ok(len(F1.worksheet("NGUOI_DUNG").values) == 1 and len(F2.worksheet("GIAO_DICH").values) == 1,
       "thêm hàng không ghi phiếu vào file nhà máy nào")
    t.append("NGUOI_DUNG", [dict(ma="K1", ten="K1", nha_may="NM1", quyen="KHO", pin_hash="x", hoat_dong="1", cap_nhat_luc="2026-11-27 08:00:00"),
                            dict(ma="K2", ten="K2", nha_may="NM2", quyen="KHO", pin_hash="x", hoat_dong="1", cap_nhat_luc="2026-11-27 08:00:00")])
    ok(len(F1.worksheet("NGUOI_DUNG").values) == 2 and len(F2.worksheet("NGUOI_DUNG").values) == 2
       and F1.worksheet("NGUOI_DUNG").values[1][0] == "K1" and F2.worksheet("NGUOI_DUNG").values[1][0] == "K2",
       "người dùng NM1 vào file NM1, người dùng NM2 vào file NM2")
    ok(set(t.read_all()["NGUOI_DUNG"]["ma"]) == {"K1", "K2"}, "đọc lại: đăng nhập thấy người dùng của cả hai file")

    hn = str(L.hom_nay())
    sp2 = L.lap_phieu(t, t.read_all(), K_2, P("N1", hn, "NM2", [Dg(mh, 10, kg_hoa_don=733.0)], doi_tac="NCC A"))
    sp1 = L.lap_phieu(t, t.read_all(), K_1, P("N1", hn, "NM1", [Dg(mh, 4)], doi_tac="NCC B"))
    ok(len(F2.worksheet("GIAO_DICH").values) == 2 and len(F1.worksheet("GIAO_DICH").values) == 2,
       "phiếu NM2 chỉ vào file NM2, phiếu NM1 chỉ vào file NM1")
    ok(F2.worksheet("GIAO_DICH").values[1][1] == sp2 and F1.worksheet("GIAO_DICH").values[1][1] == sp1, "mỗi file giữ đúng số phiếu của nhà máy mình")
    ok(len(DM3.worksheet("DOI_TAC").values) == 3, "đối tác tự thêm vào danh mục CHUNG (cả hai nhà máy cùng thấy)")
    d3 = t.read_all()
    ok(len(d3["GIAO_DICH"]) == 2 and L.ton_mot(d3, "NM2", mh) == 10 and L.ton_mot(d3, "NM1", mh) == 4,
       "đọc gộp hai file; tồn từng nhà máy tách riêng (NM2 = 10, NM1 = 4)")
    ok(len(F1.worksheet("NHAT_KY").values) >= 3 and len(F2.worksheet("NHAT_KY").values) >= 3, "nhật ký ghi vào file của nhà máy người thao tác")
    # chuyển kho NM2 -> NM1 đi qua hai file
    x5 = L.lap_phieu(t, t.read_all(), K_2, P("X5", hn, "NM2", [Dg(mh, 3)], kho_den="NM1"))
    cho = L.x5_cho_nhan(t.read_all(), "NM1")
    ok(len(cho) == 1 and len(F2.worksheet("GIAO_DICH").values) == 3, "X5 ghi ở file NM2, NM1 thấy phiếu chờ nhận")
    n4 = L.phieu_nhan_tu_x5(t.read_all(), cho.iloc[0]["phieu_id"], hn)
    L.lap_phieu(t, t.read_all(), K_1, n4)
    d3 = t.read_all()
    ok(len(F1.worksheet("GIAO_DICH").values) == 3 and L.ton_mot(d3, "NM1", mh) == 7 and L.ton_mot(d3, "NM2", mh) == 7,
       "N4 ghi ở file NM1; tồn sau chuyển kho: NM2 giảm 3 (=7), NM1 tăng 3 (=7)")
    # hủy phiếu: dòng HUY_PHIEU vào file của nhà máy có phiếu
    L.huy_phieu(t, t.read_all(), K_1, d3["GIAO_DICH"][d3["GIAO_DICH"]["so_phieu"] == sp1].iloc[0]["phieu_id"], "thử hủy")
    ok(len(F1.worksheet("HUY_PHIEU").values) == 2 and len(F2.worksheet("HUY_PHIEU").values) == 1,
       "hủy phiếu NM1 -> HUY_PHIEU của file NM1")
    L.khoa_so(t, t.read_all(), QT1, "NM2", str(L.hom_nay() - __import__("datetime").timedelta(days=1)))
    ok(len(F2.worksheet("KY_KHOA").values) == 2 and len(F1.worksheet("KY_KHOA").values) == 1, "khóa sổ NM2 -> KY_KHOA của file NM2")
    try:
        t.append("HUY_PHIEU", [dict(phieu_id="x", so_phieu="XX-N1-1", ly_do="", nguoi_huy="a", thoi_diem="")])
        ok(False, "phải báo không biết ghi vào file nào")
    except ValueError as e:
        ok("thiếu nha_may" in str(e), "dòng không rõ thuộc nhà máy nào -> từ chối ghi, không đoán")

    # ---- Script chuyển từ 1 file phát sinh chung sang 2 file nhà máy ----
    from scripts.chuyen_tach_nha_may import chuyen
    O_DM = FakeSS("APP_KHO_DANH_MUC", "https://docs.google.com/spreadsheets/d/ODM")
    O_PS = FakeSS("APP_KHO_PHAT_SINH_2026", "https://docs.google.com/spreadsheets/d/OPS")
    M1 = FakeSS("APP_KHO_PHAT_SINH_NM1_2026", "https://docs.google.com/spreadsheets/d/M1")
    M2 = FakeSS("APP_KHO_PHAT_SINH_NM2_2026", "https://docs.google.com/spreadsheets/d/M2")
    cai_gspread({"ODM": O_DM, "OPS": O_PS, "M1": M1, "M2": M2})
    cu = mo_store(dict(kieu="sheets", sheet_id_danh_muc="ODM", sheet_id_phat_sinh="OPS", credentials={"x": 1}))
    cu.khoi_tao()
    cu.append("KHO_MISA", [dict(nha_may="NM2", so="MTL", kho_misa="152.SX"), dict(nha_may="NM1", so="MTL", kho_misa="152.GC")])
    cu.append("MA_MISA", [dict(so="MTL", ma_misa="THEPTAM004", ten_misa="Thép tấm 3 ly", dvt_misa="kg")])
    mc = L.them_hang(cu, cu.read_all(), QT1, so="MTL", ma_misa="THEPTAM004", ten_hang="Thép tấm 3 x 1220 x 2500mm", nhom="Thép tấm",
                     dvt_kho="tấm", cach="KICH_THUOC", hinh_dang="tấm", vat_lieu="thép", kt=[3, 1220, 2500])
    cu.append("NGUOI_DUNG", [dict(ma="K1", ten="K1", nha_may="NM1", quyen="KHO", pin_hash="x", hoat_dong="1", cap_nhat_luc="t"),
                             dict(ma="K2", ten="K2", nha_may="NM2", quyen="KHO", pin_hash="x", hoat_dong="1", cap_nhat_luc="t")])
    for nm in NHA_MAY:
        pk = P("K1", "2026-11-23", nm, [Dg(mc, 100 if nm == "NM1" else 250, dau=1)], ly_do="Tồn đầu")
        L.lap_phieu(cu, cu.read_all(), QT1, pk)
    L.lap_phieu(cu, cu.read_all(), K_2, P("N1", hn, "NM2", [Dg(mc, 5)], doi_tac="NCC"))   # phiếu thử
    moi = mo_store(dict(kieu="sheets", sheet_id_danh_muc="ODM", sheet_id_nm1="M1", sheet_id_nm2="M2", credentials={"x": 1}))
    bao = chuyen(cu, moi, thuc_hien=False)
    ok(any("CHẠY THỬ" in b for b in bao) and len(M1.worksheets()) == 1 and len(M2.worksheets()) == 1,
       "chạy thử: chỉ báo cáo, không tạo sheet / không ghi gì vào file mới")
    bao = chuyen(cu, moi, thuc_hien=True)
    ok(any("đối chiếu ĐẠT" in b for b in bao), "chuyển thật: đối chiếu đạt")
    dn = moi.read_all()
    g = dn["GIAO_DICH"]
    ok(set(g["loai"]) == {"K1"} and len(g) == 2, "mặc định chỉ chuyển tồn đầu K1 (phiếu thử N1 không chuyển)")
    ok(L.ton_mot(dn, "NM1", mc) == 100 and L.ton_mot(dn, "NM2", mc) == 250, "tồn đầu từng nhà máy giữ nguyên sau khi chuyển")
    ok(len(M1.worksheet("GIAO_DICH").values) == 2 and len(M2.worksheet("GIAO_DICH").values) == 2
       and set(dn["NGUOI_DUNG"]["ma"]) == {"K1", "K2"}, "mỗi nhà máy một file; người dùng chuyển theo nhà máy")
    ok(len(O_PS.worksheet("GIAO_DICH").values) == 4, "file cũ giữ nguyên (không sửa, không xóa)")
    try:
        chuyen(cu, moi, thuc_hien=True)
        ok(False, "phải từ chối chạy lần 2")
    except SystemExit as e:
        ok("đã có dữ liệu" in str(e), "chạy lần 2 bị từ chối, không ghi trùng")
    print(f"\nTẤT CẢ {N} KIỂM TRA GIẢ LẬP GOOGLE ĐẠT")


if __name__ == "__main__":
    main()
