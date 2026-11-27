"""3 báo cáo Excel cho kế toán: Báo cáo nhập, Báo cáo xuất, Báo cáo nhập – xuất – tồn.

Mỗi file tách sheet theo sổ (MTL, CMA). Phiếu đã hủy KHÔNG tính.
"""
from __future__ import annotations

import io

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .logic import bay_gio, giao_dich_hieu_luc
from .schema import DUNG_SAI, LOAI

COT_CHI_TIET = [
    ("ngay", "Ngày"), ("so_phieu", "Số phiếu"), ("loai", "Mã NV"), ("ten_loai", "Nghiệp vụ"),
    ("nha_may", "Nhà máy"), ("so", "Sổ"), ("kho_misa", "Kho MISA"), ("ma_misa", "Mã kế toán (MISA)"),
    ("ma_hang", "Mã quy cách (kho)"), ("ten_hang", "Tên hàng (kho)"), ("dvt_kho", "ĐVT kho"),
    ("sl_kho", "SL theo ĐVT kho"), ("he_so", "Hệ số (ĐVT MISA / 1 ĐVT kho)"), ("dvt_misa", "ĐVT MISA"),
    ("sl_misa", "SL quy đổi (ĐVT MISA)"), ("kg_hoa_don", "Kg theo hóa đơn"),
    ("lech_pct", "Lệch % kg hóa đơn so với quy đổi"), ("lsx", "Lệnh SX"), ("po_kh", "PO khách hàng"),
    ("doi_tac", "Đối tác"), ("bo_phan", "Tổ / bộ phận"), ("so_chung_tu", "Số chứng từ"),
    ("kho_den", "Chuyển đến"), ("ly_do", "Lý do"), ("lo", "Lô"), ("ghi_chu_dong", "Ghi chú dòng"),
    ("ghi_chu_phieu", "Ghi chú phiếu"), ("nguoi_lap", "Người lập"), ("thoi_diem_lap", "Thời điểm lập"),
]
SO_COT = {"sl_kho", "he_so", "sl_misa", "kg_hoa_don", "lech_pct"}
_VANG = PatternFill("solid", fgColor="FFE699")

_HDR = PatternFill("solid", fgColor="DDEBF7")
_THIN = Side(style="thin", color="999999")


def _loc(data, nha_may, so):
    g = giao_dich_hieu_luc(data)
    if nha_may:
        g = g[g["nha_may"] == nha_may]
    if so:
        g = g[g["so"] == so]
    return g


def _ghi_bang(ws, tieu_de: list[str], df: pd.DataFrame, cot: list[tuple[str, str]], so_cot: set[str]):
    for t in tieu_de:
        ws.append([t])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    hdr_row = len(tieu_de) + 2   # tiêu đề + 1 dòng trống (ws.max_row không đếm dòng trống nên không dùng được)
    ws.append([c[1] for c in cot])
    for cell in ws[hdr_row]:
        cell.font = Font(bold=True)
        cell.fill = _HDR
        cell.alignment = Alignment(wrap_text=True, vertical="center")
        cell.border = Border(top=_THIN, bottom=_THIN, left=_THIN, right=_THIN)
    dong = hdr_row   # tự đếm dòng: ws.max_row / ws[...] tính lại toàn bảng mỗi lần nên rất chậm với bảng lớn
    for rec in df.to_dict("records"):
        vals = []
        for k, _n in cot:
            v = rec.get(k, "")
            if k in so_cot:
                v = None if pd.isna(v) else float(v)
            else:
                v = "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
            vals.append(v)
        ws.append(vals)
        dong += 1
        for j, v in enumerate(vals, 1):
            if isinstance(v, str) and v.startswith("="):   # chuỗi bắt đầu '=' giữ là chữ, không thành công thức
                ws.cell(row=dong, column=j).data_type = "s"
            elif cot[j - 1][0] == "lech_pct" and v is not None and abs(v) > DUNG_SAI * 100 + 1e-9:
                ws.cell(row=dong, column=j).fill = _VANG   # lệch kg hóa đơn vượt ±2% -> tô vàng
    for j, (k, n) in enumerate(cot, 1):
        L = get_column_letter(j)
        ws.column_dimensions[L].width = min(max(len(n), 10) + 2, 40)
        if k in so_cot:
            for row in ws.iter_rows(min_row=hdr_row + 1, min_col=j, max_col=j):
                row[0].number_format = "#,##0.###"
    ws.freeze_panes = ws.cell(row=hdr_row + 1, column=1)
    if df.shape[0]:
        ws.auto_filter.ref = f"A{hdr_row}:{get_column_letter(len(cot))}{dong}"


def _xlsx(wb) -> bytes:
    bio = io.BytesIO()
    wb.save(bio)
    return bio.getvalue()


def _them_lech(g: pd.DataFrame) -> pd.DataFrame:
    """% lệch của kg hóa đơn so với kg quy đổi (chỉ dòng có kg hóa đơn và ĐVT MISA là kg)."""
    g = g.copy()
    ok = (g["kg_hoa_don"] > 0) & (g["sl_misa"] > 0) & (g["dvt_misa"].str.lower() == "kg")
    g["lech_pct"] = float("nan")
    g.loc[ok, "lech_pct"] = (g.loc[ok, "kg_hoa_don"] - g.loc[ok, "sl_misa"]) / g.loc[ok, "sl_misa"] * 100
    return g


def _bao_cao_chieu(data, tu, den, nha_may, dau, ten_bc) -> bytes:
    g = _loc(data, nha_may, None)
    g = g[(g["ngay"] >= tu) & (g["ngay"] <= den) & (g["dau"] == dau)]
    g = _them_lech(g).sort_values(["ngay", "so_phieu", "dong"])
    wb = Workbook()
    wb.remove(wb.active)
    for so in ["MTL", "CMA"]:
        ws = wb.create_sheet(so)
        _ghi_bang(ws, [f"{ten_bc} – sổ {so}",
                       f"Từ {tu} đến {den} – Nhà máy: {nha_may or 'NM1 + NM2'}",
                       f"Lập lúc {bay_gio():%Y-%m-%d %H:%M}. Không gồm phiếu đã hủy. "
                       f"Ô vàng = kg hóa đơn lệch quá ±{DUNG_SAI * 100:g}% so với kg quy đổi."],
                  g[g["so"] == so], COT_CHI_TIET, SO_COT)
    ten_misa = data["MA_MISA"].drop_duplicates(["so", "ma_misa"], keep="last").set_index(["so", "ma_misa"])["ten_misa"]
    for so in ["MTL", "CMA"]:   # tổng hợp theo mã: gộp mọi quy cách cùng mã để kế toán đối chiếu
        gs = g[g["so"] == so]
        tg = gs.groupby(["kho_misa", "ma_misa", "dvt_misa"], as_index=False).agg(
            so_dong=("so_phieu", "size"), sl_misa=("sl_misa", "sum"), kg_hoa_don=("kg_hoa_don", "sum"))
        tg["ten_misa"] = [ten_misa.get((so, m), "") for m in tg["ma_misa"]]
        cot = [("kho_misa", "Kho MISA"), ("ma_misa", "Mã kế toán"), ("ten_misa", "Tên hàng MISA"), ("dvt_misa", "ĐVT MISA"),
               ("so_dong", "Số dòng phiếu"), ("sl_misa", "SL quy đổi (ĐVT MISA)")]
        if dau > 0:
            cot.append(("kg_hoa_don", "Kg theo hóa đơn (cộng các dòng có ghi)"))
        ws = wb.create_sheet(f"{so}_TONG_HOP_THEO_MA")
        _ghi_bang(ws, [f"{ten_bc} – TỔNG HỢP THEO MÃ KẾ TOÁN – sổ {so}",
                       f"Từ {tu} đến {den} – Nhà máy: {nha_may or 'NM1 + NM2'}",
                       "Mọi quy cách cùng một mã được cộng lại theo ĐVT MISA (ví dụ các khổ thép 3 ly cùng vào THEPTAM004)."],
                  tg.sort_values(["kho_misa", "ma_misa"]), cot, {"so_dong", "sl_misa", "kg_hoa_don"})
    return _xlsx(wb)


def bao_cao_nhap(data, tu, den, nha_may=None) -> bytes:
    return _bao_cao_chieu(data, tu, den, nha_may, 1, "BÁO CÁO NHẬP KHO")


def bao_cao_xuat(data, tu, den, nha_may=None) -> bytes:
    return _bao_cao_chieu(data, tu, den, nha_may, -1, "BÁO CÁO XUẤT KHO")


def _nxt(g: pd.DataFrame, tu, den, khoa: list[str], cot_sl: str) -> pd.DataFrame:
    g = g.assign(gt=g["dau"] * g[cot_sl])
    dau_ky = g[g["ngay"] < tu].groupby(khoa)["gt"].sum().rename("ton_dau")
    trong = g[(g["ngay"] >= tu) & (g["ngay"] <= den)]
    out = pd.DataFrame(dau_ky)
    cols_n, cols_x = [], []
    for ma in LOAI:
        for d, ds in ((1, cols_n), (-1, cols_x)):
            s = trong[(trong["loai"] == ma) & (trong["dau"] == d)].groupby(khoa)[cot_sl].sum()
            if s.empty:
                continue
            ten = f"{'Nhập' if d > 0 else 'Xuất'} {ma}"
            out = out.join(s.rename(ten), how="outer")
            ds.append(ten)
    out = out.fillna(0.0)
    out["tong_nhap"] = out[cols_n].sum(axis=1) if cols_n else 0.0
    out["tong_xuat"] = out[cols_x].sum(axis=1) if cols_x else 0.0
    out["ton_cuoi"] = out["ton_dau"] + out["tong_nhap"] - out["tong_xuat"]
    out = out[(out[["ton_dau", "tong_nhap", "tong_xuat", "ton_cuoi"]].abs() > 1e-9).any(axis=1)]
    return out.reset_index(), cols_n, cols_x


def _them_doi_chieu(t: pd.DataFrame, data, g_all: pd.DataFrame, so: str, nha_may):
    """Thêm cột đối chiếu với số tồn MISA (sheet TON_MISA_THAM_KHAO) tại ngày chụp gần nhất.
    Thêm cả mã MISA có tồn mà app chưa có hoạt động, để thấy đủ 'MISA có – app không có'."""
    tm = data["TON_MISA_THAM_KHAO"]
    tm = tm[tm["so"] == so]
    if tm.empty:
        return t, [], None
    snap = tm["ngay"].max()
    tm = tm[tm["ngay"] == snap]
    if nha_may:
        km = data["KHO_MISA"]
        tm = tm[tm["kho_misa"].isin(set(km[km["nha_may"] == nha_may]["kho_misa"]))]
    tm = tm.assign(dvt_l=tm["dvt_misa"].str.lower()).groupby(["kho_misa", "ma_misa", "dvt_l"], as_index=False).agg(
        sl_misa=("sl_misa", "sum"), dvt_misa=("dvt_misa", "first"))
    misa = {(r.kho_misa, r.ma_misa, r.dvt_l): r.sl_misa for r in tm.itertuples()}
    gs = g_all[(g_all["so"] == so) & (g_all["ngay"] <= snap)]
    app = (gs["dau"] * gs["sl_misa"]).groupby([gs["kho_misa"], gs["ma_misa"], gs["dvt_misa"].str.lower()]).sum().to_dict()
    t = t.copy()
    khoa = lambda df: list(zip(df["kho_misa"], df["ma_misa"], df["dvt_misa"].str.lower()))
    co = set(khoa(t))
    them = [r for r in tm.itertuples() if (r.kho_misa, r.ma_misa, r.dvt_l) not in co]
    if them:
        moi = pd.DataFrame([dict(kho_misa=r.kho_misa, ma_misa=r.ma_misa, dvt_misa=r.dvt_misa) for r in them])
        for c in t.columns:
            if c not in moi:
                moi[c] = 0.0
        t = pd.concat([t, moi[t.columns]], ignore_index=True)
    k = khoa(t)
    t["misa_snap"] = [misa.get(x) for x in k]
    t["app_snap"] = [app.get(x, 0.0) for x in k]
    t["chenh_snap"] = [a - m if m is not None else None for a, m in zip(t["app_snap"], t["misa_snap"])]
    t[["misa_snap", "chenh_snap"]] = t[["misa_snap", "chenh_snap"]].astype(float)
    nhan = pd.Timestamp(snap).strftime("%d/%m/%Y") if snap else ""
    return t, [("misa_snap", f"Tồn MISA {nhan} (tham khảo)"), ("app_snap", f"Tồn app {nhan} (quy đổi)"),
               ("chenh_snap", "Chênh app − MISA")], snap


def bao_cao_nxt(data, tu, den, nha_may=None) -> bytes:
    g_all = _loc(data, nha_may, None)
    g = g_all[g_all["ngay"] <= den]
    ten_misa = data["MA_MISA"].drop_duplicates(["so", "ma_misa"], keep="last").set_index(["so", "ma_misa"])["ten_misa"]
    wb = Workbook()
    wb.remove(wb.active)
    ghi_chu = {f"{c} {k}": v["ten"] for k, v in LOAI.items() for c in ("Nhập", "Xuất")}
    for so in ["MTL", "CMA"]:
        gs = g[g["so"] == so]
        # 1) Theo mã MISA (ĐVT MISA) – để kế toán so với MISA
        khoa = ["kho_misa", "ma_misa", "dvt_misa"]
        t, cn, cx = _nxt(gs, tu, den, khoa, "sl_misa")
        t, cot_dc, snap = _them_doi_chieu(t, data, g_all, so, nha_may)
        t.insert(2, "ten_misa", [ten_misa.get((so, m), "") for m in t["ma_misa"]])
        cot = [("kho_misa", "Kho MISA"), ("ma_misa", "Mã kế toán"), ("ten_misa", "Tên hàng MISA"),
               ("dvt_misa", "ĐVT MISA"), ("ton_dau", "Tồn đầu kỳ")] + [(c, c) for c in cn] + \
              [("tong_nhap", "Tổng nhập")] + [(c, c) for c in cx] + [("tong_xuat", "Tổng xuất"),
                                                                      ("ton_cuoi", "Tồn cuối kỳ")] + cot_dc
        ws = wb.create_sheet(f"{so}_THEO_MA_MISA")
        tieu = [f"BÁO CÁO NHẬP – XUẤT – TỒN theo mã kế toán – sổ {so} (ĐVT MISA)",
                f"Từ {tu} đến {den} – Nhà máy: {nha_may or 'NM1 + NM2'}",
                "Mã NV: " + "; ".join(f"{c} = {ghi_chu.get(c, c)}" for c in cn + cx)]
        if snap:
            tieu.append(f"Ba cột cuối chỉ để đối chiếu: số MISA ngày {snap} do kế toán cung cấp; “Tồn app” tính theo hệ số quy đổi; "
                        f"khác nhau là chênh lệch cần kiểm kê. Không dùng để hạch toán.")
        _ghi_bang(ws, tieu, t.sort_values(["kho_misa", "ma_misa"]), cot,
                  {"ton_dau", "tong_nhap", "tong_xuat", "ton_cuoi", *cn, *cx, "misa_snap", "app_snap", "chenh_snap"})
        # 2) Theo quy cách kho (ĐVT kho)
        khoa = ["nha_may", "ma_hang", "ten_hang", "dvt_kho", "ma_misa"]
        t, cn, cx = _nxt(gs, tu, den, khoa, "sl_kho")
        cot = [("nha_may", "Nhà máy"), ("ma_misa", "Mã kế toán"), ("ma_hang", "Mã quy cách"),
               ("ten_hang", "Tên hàng (kho)"), ("dvt_kho", "ĐVT kho"), ("ton_dau", "Tồn đầu kỳ")] + \
              [(c, c) for c in cn] + [("tong_nhap", "Tổng nhập")] + [(c, c) for c in cx] + \
              [("tong_xuat", "Tổng xuất"), ("ton_cuoi", "Tồn cuối kỳ")]
        ws = wb.create_sheet(f"{so}_THEO_QUY_CACH")
        _ghi_bang(ws, [f"BÁO CÁO NHẬP – XUẤT – TỒN theo quy cách kho – sổ {so} (ĐVT kho)",
                       f"Từ {tu} đến {den} – Nhà máy: {nha_may or 'NM1 + NM2'}",
                       "Mã NV: " + "; ".join(f"{c} = {ghi_chu.get(c, c)}" for c in cn + cx)],
                  t.sort_values(["nha_may", "ma_misa", "ten_hang"]), cot,
                  {"ton_dau", "tong_nhap", "tong_xuat", "ton_cuoi", *cn, *cx})
    return _xlsx(wb)
