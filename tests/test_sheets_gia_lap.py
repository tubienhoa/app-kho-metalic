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
    print(f"\nTẤT CẢ {N} KIỂM TRA GIẢ LẬP GOOGLE ĐẠT")


if __name__ == "__main__":
    main()
