"""Nghiệp vụ: lập phiếu, hủy phiếu, tính tồn, kiểm kê, khóa sổ.

Nguyên tắc:
- Tồn KHÔNG lưu sẵn: tồn = tổng (dau * sl) của các dòng GIAO_DICH chưa bị hủy.
- Mỗi phiếu ghi 1 lần (1 lệnh append) vào GIAO_DICH; có phieu_id để không ghi trùng khi bấm lại.
- Hệ số quy đổi lấy dòng QUY_DOI đã duyệt, hiệu lực gần nhất <= ngày phiếu, và lưu cố định vào dòng phiếu.
"""
from __future__ import annotations

import unicodedata
import uuid
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone

import pandas as pd

from . import quydoi
from .schema import DUNG_SAI, LOAI, NHA_MAY, NHAN_O, QUYEN_QUAN_TRI

VN_TZ = timezone(timedelta(hours=7))
EPS = 1e-9


class LoiNghiepVu(Exception):
    """Lỗi người dùng cần sửa (thông báo tiếng Việt dễ hiểu)."""


def bay_gio() -> datetime:
    return datetime.now(VN_TZ)


def hom_nay() -> date:
    return bay_gio().date()


def bo_dau(s: str) -> str:
    s = unicodedata.normalize("NFD", str(s)).replace("đ", "d").replace("Đ", "D")
    return "".join(c for c in s if unicodedata.category(c) != "Mn").lower()


def tim_hang(hang: pd.DataFrame, tu_khoa: str) -> pd.DataFrame:
    """Tìm theo tên/mã, không cần dấu, mọi từ đều phải có."""
    if not tu_khoa.strip():
        return hang
    tu = bo_dau(tu_khoa).split()
    chuoi = (hang["ma_hang"] + " " + hang["ten_hang"] + " " + hang["ma_misa"]).map(bo_dau)
    mask = pd.Series(True, index=hang.index)
    for t in tu:
        mask &= chuoi.str.contains(t, regex=False)
    return hang[mask]


def doc_ngay(text: str, nam_mac_dinh: int | None = None) -> date | None:
    """Đọc ngày người dùng gõ: 12/11/2026, 12-11-2026, 12.11.2026, 12/11 (thiếu năm -> năm mặc định).
    Trả về None nếu không đọc được (sai định dạng hoặc ngày không có thật, vd 31/02)."""
    import re
    t = str(text).strip()
    m = re.fullmatch(r"(\d{1,2})\s*[/.\-]\s*(\d{1,2})(?:\s*[/.\-]\s*(\d{4}))?", t)
    if not m:
        return None
    ngay, thang = int(m.group(1)), int(m.group(2))
    nam = int(m.group(3)) if m.group(3) else (nam_mac_dinh or hom_nay().year)
    try:
        return date(nam, thang, ngay)
    except ValueError:
        return None


def _co(v) -> bool:
    return str(v).strip().lower() not in ("", "0", "khong", "không", "false", "n", "no")


# ---------------- Danh mục ----------------

def hang_hoat_dong(data) -> pd.DataFrame:
    h = data["HANG_KHO"].drop_duplicates("ma_hang", keep="last")
    return h[h["hoat_dong"].map(_co)].reset_index(drop=True)


def hang_cua_nha_may(data, nha_may: str) -> pd.DataFrame:
    """Hàng đang dùng mà nhà máy này được thấy: đã có phát sinh/tồn đầu ở nhà máy này, hoặc được quản trị 'mở' (HANG_NHA_MAY).
    Danh mục vẫn là một bản chung; đây chỉ là bộ lọc để nhân viên không thấy hàng của nhà máy khác."""
    h = hang_hoat_dong(data)
    gd = data["GIAO_DICH"]
    co = set(gd[gd["nha_may"] == nha_may]["ma_hang"])
    mo = data.get("HANG_NHA_MAY")   # chưa có sheet này (file/bộ nhớ đọc theo bản cũ) thì coi như chưa mở hàng nào
    if mo is not None and not mo.empty:
        co |= set(mo[mo["nha_may"] == nha_may]["ma_hang"])
    return h[h["ma_hang"].isin(co)].reset_index(drop=True)


def mo_hang_cho_nha_may(store, data, user: dict, ma_hang: str, nha_may: str):
    """Quản trị cho phép nhà máy `nha_may` nhập/xuất hàng này (chưa từng có phát sinh ở đó)."""
    _can_qt(user)
    if nha_may not in NHA_MAY:
        raise LoiNghiepVu("Nhà máy không đúng.")
    h = hang_hoat_dong(data)
    if ma_hang not in set(h["ma_hang"]):
        raise LoiNghiepVu("Không tìm thấy hàng (hoặc hàng đã ngừng dùng).")
    if ma_hang in set(hang_cua_nha_may(data, nha_may)["ma_hang"]):
        raise LoiNghiepVu(f"Hàng này đã được mở cho {nha_may}.")
    store.append("HANG_NHA_MAY", [dict(nha_may=nha_may, ma_hang=ma_hang, nguoi=user["ma"],
                                       thoi_diem=bay_gio().strftime("%Y-%m-%d %H:%M:%S"))])
    ghi_nhat_ky(store, user["ma"], "MO_HANG", f"{ma_hang} cho {nha_may}", user["nha_may"])


def nguoi_dung(data) -> pd.DataFrame:
    """Mỗi mã 1 dòng: dòng cập nhật MỚI NHẤT (theo cap_nhat_luc; dữ liệu 2 nhà máy được gộp nên không dựa vào thứ tự dòng)."""
    nd = data["NGUOI_DUNG"].sort_values("cap_nhat_luc", kind="stable")
    return nd.drop_duplicates("ma", keep="last").reset_index(drop=True)


def he_so(data, ma_hang: str, ngay: str):
    """Dòng QUY_DOI đã duyệt, hiệu lực gần nhất <= ngay. None nếu chưa có."""
    q = data["QUY_DOI"]
    q = q[(q["ma_hang"] == ma_hang) & (q["nguoi_duyet"] != "") & (q["he_so"] > 0)]
    q = q[(q["hieu_luc_tu"] == "") | (q["hieu_luc_tu"] <= ngay)]
    if q.empty:
        return None
    q = q.assign(_k=q["hieu_luc_tu"].replace("", "0000-00-00"))
    return q.sort_values("_k", kind="stable").iloc[-1]


def he_so_bang(data, ngay: str) -> pd.DataFrame:
    """Dòng hệ số hiệu lực tại `ngay` của MỌI hàng (index = ma_hang). Cùng quy tắc với he_so() nhưng tính 1 lần."""
    q = data["QUY_DOI"]
    q = q[(q["nguoi_duyet"] != "") & (q["he_so"] > 0)]
    q = q[(q["hieu_luc_tu"] == "") | (q["hieu_luc_tu"] <= ngay)]
    q = q.assign(_k=q["hieu_luc_tu"].replace("", "0000-00-00")).sort_values("_k", kind="stable")
    return q.drop_duplicates("ma_hang", keep="last").set_index("ma_hang")


def kho_misa(data, nha_may: str, so: str) -> str:
    k = data["KHO_MISA"]
    k = k[(k["nha_may"] == nha_may) & (k["so"] == so)]
    return k.iloc[-1]["kho_misa"] if not k.empty else ""


def so_cua_nha_may(data, nha_may: str) -> list[str]:
    k = data["KHO_MISA"]
    return list(dict.fromkeys(k[k["nha_may"] == nha_may]["so"]))


def khoa_den(data, nha_may: str) -> str:
    k = data["KY_KHOA"]
    k = k[k["nha_may"] == nha_may]
    return k["khoa_den_ngay"].max() if not k.empty else ""


# ---------------- Giao dịch & tồn ----------------

def giao_dich_hieu_luc(data) -> pd.DataFrame:
    gd = data["GIAO_DICH"]
    huy = set(data["HUY_PHIEU"]["phieu_id"])
    return gd[~gd["phieu_id"].isin(huy)]


def tinh_ton(data, nha_may: str | None = None, den_ngay: str | None = None,
             gd: pd.DataFrame | None = None) -> pd.DataFrame:
    """Tồn theo (nha_may, ma_hang). den_ngay: tính các phiếu có ngày <= den_ngay."""
    g = giao_dich_hieu_luc(data) if gd is None else gd
    if nha_may:
        g = g[g["nha_may"] == nha_may]
    if den_ngay:
        g = g[g["ngay"] <= den_ngay]
    g = g.assign(ton_kho=g["dau"] * g["sl_kho"], ton_misa=g["dau"] * g["sl_misa"])
    t = g.groupby(["nha_may", "ma_hang"], as_index=False)[["ton_kho", "ton_misa"]].sum()
    return t


def ton_mot(data, nha_may: str, ma_hang: str) -> float:
    t = tinh_ton(data, nha_may)
    r = t[t["ma_hang"] == ma_hang]
    return float(r["ton_kho"].iloc[0]) if not r.empty else 0.0


def phieu_tong_hop(data) -> pd.DataFrame:
    """Mỗi phiếu 1 dòng (đầu phiếu) + trạng thái hủy."""
    gd = data["GIAO_DICH"]
    if gd.empty:
        return pd.DataFrame(columns=["phieu_id", "so_phieu", "loai", "ten_loai", "ngay", "nha_may",
                                     "nguoi_lap", "thoi_diem_lap", "so_dong", "da_huy"])
    dau = gd.drop_duplicates("phieu_id").copy()
    dau["so_dong"] = dau["phieu_id"].map(gd.groupby("phieu_id").size())
    dau["da_huy"] = dau["phieu_id"].isin(set(data["HUY_PHIEU"]["phieu_id"]))
    return dau.sort_values("thoi_diem_lap", ascending=False).reset_index(drop=True)


def x5_cho_nhan(data, nha_may_nhan: str) -> pd.DataFrame:
    """Phiếu X5 chuyển tới nha_may_nhan, chưa hủy, chưa có N4 hiệu lực."""
    gd = giao_dich_hieu_luc(data)
    x5 = gd[(gd["loai"] == "X5") & (gd["kho_den"] == nha_may_nhan)]
    da_nhan = set(gd[gd["loai"] == "N4"]["phieu_lien_quan"])
    return x5[~x5["phieu_id"].isin(da_nhan)]


# ---------------- Lập phiếu ----------------

@dataclass
class DongNhap:
    ma_hang: str
    sl_kho: float
    dau: int = 0          # chỉ dùng cho X8/X9/K1 (+1 nhập, -1 xuất)
    lo: str = ""
    ghi_chu: str = ""
    kg_hoa_don: float = 0.0   # chỉ N1: số kg ghi trên hóa đơn (0 = không nhập)


@dataclass
class PhieuNhap:
    loai: str
    ngay: str             # 'YYYY-MM-DD'
    nha_may: str
    dong: list[DongNhap] = field(default_factory=list)
    doi_tac: str = ""
    bo_phan: str = ""
    lsx: str = ""
    po_kh: str = ""
    so_chung_tu: str = ""
    kho_den: str = ""
    phieu_lien_quan: str = ""
    ly_do: str = ""
    ghi_chu_phieu: str = ""
    phieu_id: str = field(default_factory=lambda: uuid.uuid4().hex)


def _so_phieu(data, p: PhieuNhap, ma_nguoi: str) -> str:
    ngay6 = p.ngay[2:4] + p.ngay[5:7] + p.ngay[8:10]
    tien_to = f"{p.nha_may}-{p.loai}-{ngay6}-{ma_nguoi}-"
    da_co = data["GIAO_DICH"]["so_phieu"]
    n = da_co[da_co.str.startswith(tien_to)].drop_duplicates().size + 1
    return f"{tien_to}{n:03d}"


def kiem_tra_phieu(data, user: dict, p: PhieuNhap) -> list[dict]:
    """Kiểm tra và dựng các dòng GIAO_DICH. Ném LoiNghiepVu nếu sai."""
    if p.loai not in LOAI:
        raise LoiNghiepVu(f"Loại phiếu {p.loai} không có.")
    cfg = LOAI[p.loai]
    la_qt = user["quyen"] == QUYEN_QUAN_TRI
    if p.loai == "K1" and not la_qt:
        raise LoiNghiepVu("Chỉ người quản trị được lập phiếu kiểm kê.")
    if not la_qt and p.nha_may != user["nha_may"]:
        raise LoiNghiepVu(f"Bạn chỉ được lập phiếu cho {user['nha_may']}.")
    try:
        ngay_d = date.fromisoformat(p.ngay)
    except ValueError:
        raise LoiNghiepVu("Ngày phiếu không đúng.")
    if ngay_d > hom_nay():
        raise LoiNghiepVu("Ngày phiếu không được sau hôm nay.")
    kd = khoa_den(data, p.nha_may)
    if kd and p.ngay <= kd:
        raise LoiNghiepVu(f"Sổ {p.nha_may} đã khóa đến {kd}. Không lập phiếu ngày {p.ngay} được.")

    for o in cfg["bat_buoc"]:
        if not str(getattr(p, o)).strip():
            raise LoiNghiepVu(f"Chưa điền: {NHAN_O[o]}.")
    if p.loai == "X5" and p.kho_den == p.nha_may:
        raise LoiNghiepVu("Nhà máy nhận phải khác nhà máy chuyển.")
    if not p.dong:
        raise LoiNghiepVu("Phiếu chưa có dòng hàng nào.")

    hang = hang_hoat_dong(data).set_index("ma_hang")
    rows, so_xuat, so_nhap = [], set(), set()
    for i, d in enumerate(p.dong, 1):
        if d.ma_hang not in hang.index:
            raise LoiNghiepVu(f"Dòng {i}: mã hàng {d.ma_hang} không có hoặc đã ngừng dùng.")
        h = hang.loc[d.ma_hang]
        try:
            sl = float(d.sl_kho)
        except (TypeError, ValueError):
            raise LoiNghiepVu(f"Dòng {i}: số lượng không đúng.")
        if not sl > 0:
            raise LoiNghiepVu(f"Dòng {i} ({h['ten_hang']}): số lượng phải lớn hơn 0.")
        dau = cfg["dau"] if cfg["dau"] != 0 else d.dau
        if dau not in (1, -1):
            raise LoiNghiepVu(f"Dòng {i}: chưa chọn nhập hay xuất.")
        hs = he_so(data, d.ma_hang, p.ngay)
        if hs is None:
            raise LoiNghiepVu(f"Dòng {i} ({h['ten_hang']}): chưa có hệ số quy đổi đã duyệt. Báo quản trị.")
        if hs["dvt_kho"] and hs["dvt_kho"] != h["dvt_kho"]:
            raise LoiNghiepVu(f"Dòng {i} ({h['ten_hang']}): ĐVT hệ số ({hs['dvt_kho']}) khác ĐVT kho ({h['dvt_kho']}). Báo quản trị.")
        kg_hd = d.kg_hoa_don or 0.0
        if kg_hd:
            if p.loai != "N1":
                raise LoiNghiepVu(f"Dòng {i}: “kg theo hóa đơn” chỉ ghi ở phiếu nhập mua (N1).")
            if not float(kg_hd) > 0:
                raise LoiNghiepVu(f"Dòng {i} ({h['ten_hang']}): kg theo hóa đơn phải lớn hơn 0.")
        (so_xuat if dau < 0 else so_nhap).add(h["so"])
        km = kho_misa(data, p.nha_may, h["so"])
        if not km:
            raise LoiNghiepVu(f"Dòng {i}: {p.nha_may} chưa khai kho MISA cho sổ {h['so']}. Báo quản trị.")
        rows.append(dict(dong=i, dau=dau, ma_hang=d.ma_hang, ten_hang=h["ten_hang"], so=h["so"],
                         ma_misa=h["ma_misa"], kho_misa=km, dvt_kho=h["dvt_kho"], sl_kho=sl,
                         he_so=float(hs["he_so"]), dvt_misa=hs["dvt_misa"],
                         sl_misa=round(sl * float(hs["he_so"]), 6),
                         lo=d.lo.strip(), ghi_chu_dong=d.ghi_chu.strip(),
                         kg_hoa_don=float(kg_hd) if kg_hd else ""))

    tat_ca_so = so_xuat | so_nhap
    if p.loai == "X9":
        if len(so_xuat) != 1 or len(so_nhap) != 1 or so_xuat == so_nhap:
            raise LoiNghiepVu("Chuyển sổ: các dòng XUẤT phải cùng 1 sổ, các dòng NHẬP phải là sổ còn lại.")
    elif len(tat_ca_so) > 1:
        raise LoiNghiepVu("Một phiếu chỉ được chứa hàng của MỘT sổ (MTL hoặc CMA). Hàng sổ kia lập phiếu riêng.")
    if p.loai == "X8" and not (so_xuat and so_nhap):
        raise LoiNghiepVu("Đổi quy cách phải có ít nhất 1 dòng XUẤT (hàng đem cắt) và 1 dòng NHẬP (hàng sau cắt).")

    # Tồn không được âm sau phiếu (trừ kiểm kê – quản trị đặt lại số)
    if p.loai != "K1":
        ton = tinh_ton(data, p.nha_may).set_index("ma_hang")["ton_kho"]
        thay_doi: dict[str, float] = {}
        for r in rows:
            thay_doi[r["ma_hang"]] = thay_doi.get(r["ma_hang"], 0.0) + r["dau"] * r["sl_kho"]
        for ma, delta in thay_doi.items():
            sau = float(ton.get(ma, 0.0)) + delta
            if delta < 0 and sau < -EPS:
                h = hang.loc[ma]
                raise LoiNghiepVu(f"{h['ten_hang']}: tồn {p.nha_may} chỉ còn {ton.get(ma, 0.0):g} {h['dvt_kho']}, "
                                  f"không xuất {-delta:g} được.")

    if p.loai == "N4":
        x5 = x5_cho_nhan(data, p.nha_may)
        if p.phieu_lien_quan not in set(x5["phieu_id"]):
            raise LoiNghiepVu("Phiếu chuyển kho này không còn chờ nhận (đã nhận hoặc đã hủy).")
    return rows


def lap_phieu(store, data, user: dict, p: PhieuNhap) -> str:
    """Ghi phiếu. data phải là bản đọc MỚI NHẤT. Trả về số phiếu."""
    gd = data["GIAO_DICH"]
    if p.phieu_id in set(gd["phieu_id"]):  # bấm lưu lần 2 sau khi đã ghi được
        return gd[gd["phieu_id"] == p.phieu_id]["so_phieu"].iloc[0]
    dong = kiem_tra_phieu(data, user, p)
    so_phieu = _so_phieu(data, p, user["ma"])
    now = bay_gio().strftime("%Y-%m-%d %H:%M:%S")
    dau_phieu = dict(phieu_id=p.phieu_id, so_phieu=so_phieu, loai=p.loai, ten_loai=LOAI[p.loai]["ten"],
                     ngay=p.ngay, nha_may=p.nha_may, nguoi_lap=user["ma"], thoi_diem_lap=now,
                     doi_tac=p.doi_tac.strip(), bo_phan=p.bo_phan.strip(), lsx=p.lsx.strip(),
                     po_kh=p.po_kh.strip(), so_chung_tu=p.so_chung_tu.strip(), kho_den=p.kho_den,
                     phieu_lien_quan=p.phieu_lien_quan, ly_do=p.ly_do.strip(),
                     ghi_chu_phieu=p.ghi_chu_phieu.strip())
    store.append("GIAO_DICH", [{**dau_phieu, **d} for d in dong])
    ghi_nhat_ky(store, user["ma"], "LAP_PHIEU", f"{so_phieu} ({len(dong)} dòng)", p.nha_may)
    try:   # phiếu đã ghi xong; lỗi ghi nhớ đối tác không được làm hỏng phiếu
        _ghi_nho_doi_tac(store, data, user, p, so_phieu)
    except Exception as e:
        try:   # không làm hỏng phiếu, nhưng để lại dấu vết trong nhật ký
            ghi_nhat_ky(store, user["ma"], "LOI_GHI_NHO_DOI_TAC", f"{so_phieu}: {type(e).__name__}: {e}"[:300], p.nha_may)
        except Exception:
            pass
    return so_phieu


def _ghi_nho_doi_tac(store, data, user: dict, p: "PhieuNhap", so_phieu: str):
    """Đối tác gõ tay mới (chưa có trong DOI_TAC cùng loại) được thêm vào danh sách để lần sau chọn lại."""
    loai_dt = LOAI[p.loai].get("doi_tac")
    ten = p.doi_tac.strip()
    if not loai_dt or not ten:
        return
    dt = data["DOI_TAC"]
    chuan = lambda t: " ".join(bo_dau(t).split())
    da_co = dt[dt["loai"] == loai_dt]["ten_doi_tac"].map(chuan)
    if chuan(ten) in set(da_co):
        return
    import re
    so = [int(m.group(1)) for x in dt["ma_doi_tac"] if (m := re.fullmatch(r"DT(\d+)", x))]
    store.append("DOI_TAC", [dict(ma_doi_tac=f"DT{(max(so) if so else 0) + 1:05d}", ten_doi_tac=ten, loai=loai_dt,
                                  ghi_chu=f"Tự thêm từ phiếu {so_phieu} ({user['ma']})")])


def phieu_nhan_tu_x5(data, x5_id: str, ngay: str) -> PhieuNhap:
    """Dựng phiếu N4 nhận đủ theo phiếu X5."""
    gd = data["GIAO_DICH"]
    x = gd[gd["phieu_id"] == x5_id].sort_values("dong")
    if x.empty:
        raise LoiNghiepVu("Không tìm thấy phiếu chuyển kho.")
    dau = x.iloc[0]
    return PhieuNhap(loai="N4", ngay=ngay, nha_may=dau["kho_den"], phieu_lien_quan=x5_id,
                     ghi_chu_phieu=f"Nhận theo {dau['so_phieu']}",
                     dong=[DongNhap(r["ma_hang"], float(r["sl_kho"]), lo=r["lo"]) for _, r in x.iterrows()])


# ---------------- Hủy phiếu ----------------

def huy_phieu(store, data, user: dict, phieu_id: str, ly_do: str):
    if not ly_do.strip():
        raise LoiNghiepVu("Phải ghi lý do hủy.")
    ds = phieu_tong_hop(data)
    r = ds[ds["phieu_id"] == phieu_id]
    if r.empty:
        raise LoiNghiepVu("Không tìm thấy phiếu.")
    r = r.iloc[0]
    if r["da_huy"]:
        raise LoiNghiepVu("Phiếu đã hủy trước đó.")
    la_qt = user["quyen"] == QUYEN_QUAN_TRI
    if not la_qt and not (r["nguoi_lap"] == user["ma"] and r["thoi_diem_lap"][:10] == str(hom_nay())):
        raise LoiNghiepVu("Bạn chỉ hủy được phiếu do mình lập trong hôm nay. Phiếu khác nhờ quản trị.")
    kd = khoa_den(data, r["nha_may"])
    if kd and r["ngay"] <= kd:
        raise LoiNghiepVu(f"Phiếu thuộc kỳ đã khóa (đến {kd}). Không hủy được – lập phiếu điều chỉnh.")
    if r["loai"] == "X5":
        n4 = giao_dich_hieu_luc(data)
        if phieu_id in set(n4[n4["loai"] == "N4"]["phieu_lien_quan"]):
            raise LoiNghiepVu("Nhà máy nhận đã nhận phiếu chuyển này. Hủy phiếu nhận (N4) trước.")
    # Hủy không được làm tồn âm
    gd_sau = giao_dich_hieu_luc(data)
    gd_sau = gd_sau[gd_sau["phieu_id"] != phieu_id]
    dong = data["GIAO_DICH"][data["GIAO_DICH"]["phieu_id"] == phieu_id]
    ton_sau = tinh_ton(data, r["nha_may"], gd=gd_sau).set_index("ma_hang")["ton_kho"]
    for ma in set(dong[dong["dau"] > 0]["ma_hang"]):
        if float(ton_sau.get(ma, 0.0)) < -EPS:
            ten = dong[dong["ma_hang"] == ma]["ten_hang"].iloc[0]
            raise LoiNghiepVu(f"Hủy phiếu này làm tồn {ten} bị âm (hàng đã được xuất). Không hủy được.")
    store.append("HUY_PHIEU", [dict(phieu_id=phieu_id, so_phieu=r["so_phieu"], ly_do=ly_do.strip(),
                                    nguoi_huy=user["ma"], thoi_diem=bay_gio().strftime("%Y-%m-%d %H:%M:%S"),
                                    nha_may=r["nha_may"])])   # nha_may chỉ để định tuyến file, không phải cột
    ghi_nhat_ky(store, user["ma"], "HUY_PHIEU", f"{r['so_phieu']}: {ly_do.strip()}", r["nha_may"])


# ---------------- Kiểm kê ----------------

def bang_kiem_ke(data, nha_may: str, ngay: str) -> pd.DataFrame:
    """Mẫu kiểm kê: mọi hàng đang dùng thuộc các sổ của nhà máy, kèm tồn sổ app tại ngày."""
    hang = hang_hoat_dong(data)
    hang = hang[hang["so"].isin(so_cua_nha_may(data, nha_may))]
    ton = tinh_ton(data, nha_may, ngay).set_index("ma_hang")["ton_kho"]
    df = hang[["ma_hang", "ten_hang", "nhom", "so", "ma_misa", "dvt_kho"]].copy()
    df["ton_so_app"] = df["ma_hang"].map(ton).fillna(0.0)
    df["sl_thuc_te"] = None
    df["ghi_chu"] = ""
    return df.sort_values(["so", "nhom", "ten_hang"]).reset_index(drop=True)


def phieu_kiem_ke(data, nha_may: str, ngay: str, dem: pd.DataFrame) -> tuple[list[PhieuNhap], pd.DataFrame]:
    """dem: cột ma_hang, sl_thuc_te (ô trống = không kiểm, giữ nguyên). Trả về phiếu K1 (mỗi sổ 1 phiếu) + bảng chênh."""
    dem = dem[["ma_hang", "sl_thuc_te"] + (["ghi_chu"] if "ghi_chu" in dem else [])].copy()
    dem["ma_hang"] = dem["ma_hang"].astype(str).str.strip()
    dem["sl_thuc_te"] = pd.to_numeric(dem["sl_thuc_te"], errors="coerce")
    dem = dem[dem["sl_thuc_te"].notna()]
    if (dem["sl_thuc_te"] < 0).any():
        raise LoiNghiepVu("Có số kiểm thực tế âm.")
    trung = dem[dem["ma_hang"].duplicated()]["ma_hang"].tolist()
    if trung:
        raise LoiNghiepVu(f"Mã hàng bị lặp trong file kiểm kê: {', '.join(trung[:10])}")
    hang = hang_hoat_dong(data).set_index("ma_hang")
    la = [m for m in dem["ma_hang"] if m not in hang.index]
    if la:
        raise LoiNghiepVu(f"Mã hàng không có trong danh mục: {', '.join(la[:10])}")
    ton = tinh_ton(data, nha_may, ngay).set_index("ma_hang")["ton_kho"]
    dem["ton_so_app"] = dem["ma_hang"].map(ton).fillna(0.0)
    dem["chenh_lech"] = (dem["sl_thuc_te"] - dem["ton_so_app"]).round(6)
    dem["so"] = dem["ma_hang"].map(hang["so"])
    dem["ten_hang"] = dem["ma_hang"].map(hang["ten_hang"])
    phieus = []
    for so, g in dem[dem["chenh_lech"].abs() > EPS].groupby("so"):
        p = PhieuNhap(loai="K1", ngay=ngay, nha_may=nha_may, ly_do="Kiểm kê thực tế",
                      ghi_chu_phieu=f"Kiểm kê {nha_may} {so} ngày {ngay}")
        for _, r in g.iterrows():
            p.dong.append(DongNhap(r["ma_hang"], abs(float(r["chenh_lech"])),
                                   dau=1 if r["chenh_lech"] > 0 else -1,
                                   ghi_chu=str(r.get("ghi_chu", "") or "")))
        phieus.append(p)
    return phieus, dem


# ---------------- Khóa sổ ----------------

def khoa_so(store, data, user: dict, nha_may: str, den_ngay: str):
    if user["quyen"] != QUYEN_QUAN_TRI:
        raise LoiNghiepVu("Chỉ quản trị được khóa sổ.")
    kd = khoa_den(data, nha_may)
    if kd and den_ngay <= kd:
        raise LoiNghiepVu(f"{nha_may} đã khóa đến {kd}.")
    if date.fromisoformat(den_ngay) > hom_nay():
        raise LoiNghiepVu("Không khóa sổ ngày tương lai.")
    now = bay_gio().strftime("%Y-%m-%d %H:%M:%S")
    ton = tinh_ton(data, nha_may, den_ngay)
    hang = data["HANG_KHO"].drop_duplicates("ma_hang", keep="last").set_index("ma_hang")
    gd = giao_dich_hieu_luc(data)
    dvt_misa = gd.drop_duplicates("ma_hang", keep="last").set_index("ma_hang")["dvt_misa"]
    rows = []
    for _, r in ton.iterrows():
        h = hang.loc[r["ma_hang"]] if r["ma_hang"] in hang.index else None
        rows.append(dict(nha_may=nha_may, ngay_chot=den_ngay, ma_hang=r["ma_hang"],
                         ten_hang=h["ten_hang"] if h is not None else "", so=h["so"] if h is not None else "",
                         ma_misa=h["ma_misa"] if h is not None else "",
                         dvt_kho=h["dvt_kho"] if h is not None else "", sl_kho=round(r["ton_kho"], 6),
                         dvt_misa=dvt_misa.get(r["ma_hang"], ""), sl_misa=round(r["ton_misa"], 6),
                         thoi_diem=now))
    store.append("TON_CHOT", rows)
    store.append("KY_KHOA", [dict(nha_may=nha_may, khoa_den_ngay=den_ngay, nguoi_khoa=user["ma"], thoi_diem=now)])
    ghi_nhat_ky(store, user["ma"], "KHOA_SO", f"{nha_may} đến {den_ngay}", nha_may)


# ---------------- Kiểm tra danh mục ----------------

def kiem_tra_danh_muc(data) -> pd.DataFrame:
    """Liệt kê lỗi danh mục để quản trị sửa trong Google Sheets."""
    loi = []
    hk = data["HANG_KHO"]
    for m in hk[hk["ma_hang"].duplicated(keep=False)]["ma_hang"].unique():
        loi.append(("HANG_KHO", m, "Mã hàng bị lặp (app dùng dòng cuối)"))
    misa = data["MA_MISA"]
    co_misa = set(zip(misa["so"], misa["ma_misa"]))
    bang_hs = he_so_bang(data, str(hom_nay()))
    for _, h in hang_hoat_dong(data).iterrows():
        if not h["ma_hang"]:
            loi.append(("HANG_KHO", "", "Dòng thiếu mã hàng"))
            continue
        if h["so"] not in ("MTL", "CMA"):
            loi.append(("HANG_KHO", h["ma_hang"], f"Sổ phải là MTL hoặc CMA (đang: '{h['so']}')"))
        if (h["so"], h["ma_misa"]) not in co_misa:
            loi.append(("HANG_KHO", h["ma_hang"], f"Mã MISA '{h['ma_misa']}' không có trong MA_MISA sổ {h['so']}"))
        if not h["dvt_kho"]:
            loi.append(("HANG_KHO", h["ma_hang"], "Thiếu ĐVT kho"))
        hs = bang_hs.loc[h["ma_hang"]] if h["ma_hang"] in bang_hs.index else None
        if hs is None:
            loi.append(("QUY_DOI", h["ma_hang"], "Chưa có hệ số đã duyệt → chưa lập phiếu được"))
        elif hs["dvt_kho"] != h["dvt_kho"]:
            loi.append(("QUY_DOI", h["ma_hang"], f"ĐVT kho trong hệ số ({hs['dvt_kho']}) khác HANG_KHO ({h['dvt_kho']})"))
        elif pd.notna(hs["kg_ly_thuyet"]) and hs["kg_ly_thuyet"] > 0 and str(hs["dvt_misa"]).lower() == "kg" \
                and quydoi.vuot_dung_sai(float(hs["he_so"]), float(hs["kg_ly_thuyet"])):
            d = quydoi.lech(float(hs["he_so"]), float(hs["kg_ly_thuyet"])) * 100
            loi.append(("QUY_DOI", h["ma_hang"], f"CẢNH BÁO: hệ số {hs['he_so']:g} kg lệch {d:+.1f}% so với kg lý thuyết "
                                                 f"{hs['kg_ly_thuyet']:.3f} (vượt ±{DUNG_SAI * 100:g}%)"))
    m = misa[misa["ma_misa"].duplicated(keep=False) & misa.duplicated(["so", "ma_misa"], keep=False)]
    for v in m["ma_misa"].unique():
        loi.append(("MA_MISA", v, "Mã MISA lặp trong cùng sổ"))
    return pd.DataFrame(loi, columns=["Sheet", "Mã", "Lỗi"])


def ghi_nhat_ky(store, nguoi: str, hanh_dong: str, chi_tiet: str, nha_may: str = ""):
    """Ghi nhật ký. `nha_may`: nhà máy của người thao tác (khi dữ liệu tách theo nhà máy, dòng này vào file của nhà máy đó)."""
    store.append("NHAT_KY", [dict(thoi_diem=bay_gio().strftime("%Y-%m-%d %H:%M:%S"),
                                  nguoi=nguoi, hanh_dong=hanh_dong, chi_tiet=chi_tiet, nha_may=nha_may)])


# ---------------- Kg theo hóa đơn: cảnh báo lệch quá dung sai ----------------

def lech_kg_hoa_don(sl_misa: float, dvt_misa: str, kg_hd: float):
    """Độ lệch (kg hóa đơn − kg quy đổi) / kg quy đổi. None nếu không so được (ĐVT MISA không phải kg)."""
    if str(dvt_misa).strip().lower() != "kg" or not kg_hd or kg_hd <= 0:
        return None
    return quydoi.lech(float(kg_hd), float(sl_misa))


def canh_bao_kg_hoa_don(data) -> pd.DataFrame:
    """Dòng nhập mua (N1, chưa hủy) có kg hóa đơn lệch quá ±2% so với kg quy đổi."""
    g = giao_dich_hieu_luc(data)
    g = g[(g["loai"] == "N1") & (g["kg_hoa_don"] > 0) & (g["dvt_misa"].str.lower() == "kg")].copy()
    cot = ["so_phieu", "ngay", "ten_hang", "sl_kho", "dvt_kho", "sl_misa", "kg_hoa_don", "lech_pct"]
    if g.empty:
        return pd.DataFrame(columns=cot)
    g["lech_pct"] = [quydoi.lech(a, b) * 100 if b > 0 else float("nan") for a, b in zip(g["kg_hoa_don"], g["sl_misa"])]
    g = g[g["lech_pct"].abs() > DUNG_SAI * 100 + 1e-9]
    return g[cot].sort_values("ngay", ascending=False).reset_index(drop=True)


# ---------------- Danh mục & quy đổi (CHỈ QUẢN TRỊ) ----------------

def _can_qt(user: dict):
    if user["quyen"] != QUYEN_QUAN_TRI:
        raise LoiNghiepVu("Chỉ quản trị được thêm hàng, đổi hệ số quy đổi hoặc khối lượng riêng.")


def klr_hien_hanh(data) -> dict:
    """Khối lượng riêng đang dùng: mặc định QT01, dòng mới nhất trong sheet KHOI_LUONG_RIENG ghi đè."""
    d = dict(quydoi.KLR_MAC_DINH)
    for _, r in data["KHOI_LUONG_RIENG"].iterrows():
        if r["vat_lieu"] and pd.notna(r["klr"]) and r["klr"] > 0:
            d[r["vat_lieu"]] = float(r["klr"])
    return d


def dat_klr(store, data, user: dict, vat_lieu: str, klr: float, ghi_chu: str = ""):
    _can_qt(user)
    vat_lieu = vat_lieu.strip()
    if not vat_lieu:
        raise LoiNghiepVu("Chưa ghi tên vật liệu.")
    if not (0.5 <= float(klr) <= 25):
        raise LoiNghiepVu("Khối lượng riêng phải nằm trong khoảng 0,5 – 25 g/cm³.")
    store.append("KHOI_LUONG_RIENG", [dict(vat_lieu=vat_lieu, klr=float(klr), ghi_chu=ghi_chu.strip(), nguoi=user["ma"],
                                           thoi_diem=bay_gio().strftime("%Y-%m-%d %H:%M:%S"))])
    ghi_nhat_ky(store, user["ma"], "KLR", f"{vat_lieu} = {klr:g} g/cm3", user["nha_may"])


def ma_misa_cua_so(data, so: str) -> pd.DataFrame:
    m = data["MA_MISA"]
    return m[m["so"] == so].drop_duplicates("ma_misa", keep="last").reset_index(drop=True)


def tinh_quy_doi(data, *, dvt_kho: str, dvt_misa: str, cach: str, hinh_dang: str = "", vat_lieu: str = "",
                 kt: list | None = None, kg_can: float = 0.0, he_so_nhap: float = 0.0) -> dict:
    """Tính hệ số (ĐVT MISA / 1 ĐVT kho) theo 1 trong 3 cách. Trả dict gồm he_so, kg_ly_thuyet, cach_tinh,
    klr, canh_bao (danh sách chữ). Ném LoiNghiepVu nếu thiếu/sai dữ liệu."""
    dvt_kho, dvt_misa = dvt_kho.strip().lower(), dvt_misa.strip().lower()
    out = dict(he_so=0.0, kg_ly_thuyet="", cach_tinh="", klr="", kt=[], canh_bao=[], hinh_dang="", vat_lieu="")
    if cach == "CUNG_DVT":
        if dvt_kho != dvt_misa:
            raise LoiNghiepVu(f"ĐVT kho “{dvt_kho}” khác ĐVT MISA “{dvt_misa}”: không dùng được cách “cùng đơn vị”.")
        return dict(out, he_so=1.0, cach_tinh="Cùng ĐVT")
    if cach == "TRUC_TIEP":
        if not float(he_so_nhap) > 0:
            raise LoiNghiepVu("Hệ số phải lớn hơn 0.")
        return dict(out, he_so=float(he_so_nhap), cach_tinh="Nhập trực tiếp (theo cân / chứng từ)")
    if cach != "KICH_THUOC":
        raise LoiNghiepVu("Chưa chọn cách tính hệ số.")
    if dvt_misa != "kg":
        raise LoiNghiepVu("Tính theo kích thước chỉ dùng khi ĐVT MISA là kg.")
    klr_d = klr_hien_hanh(data)
    if vat_lieu not in klr_d:
        raise LoiNghiepVu("Chưa chọn vật liệu (hoặc vật liệu chưa có khối lượng riêng).")
    try:
        lt = quydoi.kg_ly_thuyet(hinh_dang, klr_d[vat_lieu], list(kt or []))
    except ValueError as e:
        raise LoiNghiepVu(str(e))
    out.update(kg_ly_thuyet=round(lt, 6), klr=klr_d[vat_lieu], kt=[float(x) for x in kt], hinh_dang=hinh_dang,
               vat_lieu=vat_lieu)
    if kg_can and float(kg_can) > 0:
        out.update(he_so=float(kg_can),
                   cach_tinh=f"Cân thực tế {float(kg_can):g} kg/{dvt_kho} (kích thước + {vat_lieu} cho {lt:.3f} kg)")
        if quydoi.vuot_dung_sai(float(kg_can), lt):
            out["canh_bao"].append(f"kg cân {float(kg_can):g} lệch {quydoi.lech(float(kg_can), lt) * 100:+.1f}% so với kg lý "
                                   f"thuyết {lt:.3f} (vượt ±{DUNG_SAI * 100:g}%) – kiểm tra lại số cân hoặc kích thước.")
    else:
        out.update(he_so=round(lt, 6),
                   cach_tinh=f"{hinh_dang} {'×'.join(f'{x:g}' for x in kt)} mm × {vat_lieu} {klr_d[vat_lieu]:g} g/cm³ (QT01)")
    return out


def _id_hang_moi(data) -> str:
    import re
    so = [int(m.group(1)) for x in data["HANG_KHO"]["ma_hang"] if (m := re.fullmatch(r"Q(\d+)", x))]
    return f"Q{(max(so) if so else 0) + 1:05d}"


def them_hang(store, data, user: dict, *, so: str, ma_misa: str, ten_hang: str, nhom: str, dvt_kho: str,
              cach: str, ghi_chu: str = "", mo_cho: list | None = None, **kw) -> str:
    """Thêm hàng kho (1 quy cách) + dòng hệ số đầu tiên. Trả về mã quy cách mới (Qnnnnn).
    mo_cho: các nhà máy được thấy hàng này ngay (mặc định: nhà máy của quản trị đang thêm)."""
    _can_qt(user)
    ten_hang, dvt_kho, ma_misa = ten_hang.strip(), dvt_kho.strip().lower(), ma_misa.strip()
    if not ten_hang or not dvt_kho or not nhom.strip():
        raise LoiNghiepVu("Phải điền đủ tên hàng, nhóm và ĐVT kho.")
    m = ma_misa_cua_so(data, so)
    m = m[m["ma_misa"] == ma_misa]
    if m.empty:
        raise LoiNghiepVu(f"Mã MISA “{ma_misa}” không có trong danh sách mã của sổ {so}.")
    dvt_misa = m.iloc[0]["dvt_misa"]
    hk = data["HANG_KHO"]
    chuan = lambda t: " ".join(bo_dau(t).split())
    trung = hk[(hk["so"] == so) & (hk["ma_misa"] == ma_misa) & (hk["ten_hang"].map(chuan) == chuan(ten_hang))
               & (hk["dvt_kho"].str.lower() == dvt_kho)]
    if not trung.empty:
        raise LoiNghiepVu(f"Hàng này đã có trong danh mục (mã {trung.iloc[0]['ma_hang']}).")
    t = tinh_quy_doi(data, dvt_kho=dvt_kho, dvt_misa=dvt_misa, cach=cach, **kw)
    if cach == "TRUC_TIEP" and not ghi_chu.strip():
        raise LoiNghiepVu("Hệ số nhập trực tiếp phải ghi nguồn (số cân / chứng từ nào) vào ô ghi chú.")
    ma = _id_hang_moi(data)
    now = bay_gio()
    store.append("HANG_KHO", [dict(ma_hang=ma, ten_hang=ten_hang, nhom=nhom.strip(), dvt_kho=dvt_kho, so=so, ma_misa=ma_misa,
                                   vat_lieu=t["vat_lieu"], quy_cach="x".join(f"{x:g}" for x in t["kt"]), hoat_dong="1",
                                   ghi_chu=ghi_chu.strip(), nguon=f"app {user['ma']} {now:%Y-%m-%d}")])
    _ghi_dong_quy_doi(store, user, ma, dvt_kho, dvt_misa, t, "", ghi_chu, now)
    for nm in (mo_cho if mo_cho is not None else [user["nha_may"]]):
        store.append("HANG_NHA_MAY", [dict(nha_may=nm, ma_hang=ma, nguoi=user["ma"], thoi_diem=f"{now:%Y-%m-%d %H:%M:%S}")])
    ghi_nhat_ky(store, user["ma"], "THEM_HANG", f"{ma} {ten_hang} ({dvt_kho}→{dvt_misa} x{t['he_so']:g})", user["nha_may"])
    return ma


def _ghi_dong_quy_doi(store, user, ma_hang, dvt_kho, dvt_misa, t, hieu_luc_tu, ghi_chu, now):
    kt = list(t["kt"]) + [""] * (4 - len(t["kt"]))
    store.append("QUY_DOI", [dict(ma_hang=ma_hang, dvt_kho=dvt_kho, dvt_misa=dvt_misa, he_so=t["he_so"], cach_tinh=t["cach_tinh"],
                                  nguoi_duyet=user["ma"], ngay_duyet=f"{now:%Y-%m-%d}", hieu_luc_tu=hieu_luc_tu,
                                  ghi_chu=ghi_chu.strip(), hinh_dang=t["hinh_dang"], vat_lieu=t["vat_lieu"], klr=t["klr"],
                                  kt_a=kt[0], kt_b=kt[1], kt_c=kt[2], kt_d=kt[3], kg_ly_thuyet=t["kg_ly_thuyet"],
                                  kg_can=(t["he_so"] if str(t["cach_tinh"]).startswith("Cân thực tế") else ""))])


def doi_he_so(store, data, user: dict, ma_hang: str, *, cach: str, hieu_luc_tu: str, ghi_chu: str = "", **kw):
    """Thêm một dòng hệ số mới cho hàng đã có, hiệu lực từ ngày chỉ định. Phiếu đã lập giữ nguyên hệ số cũ."""
    _can_qt(user)
    hk = hang_hoat_dong(data).set_index("ma_hang")
    if ma_hang not in hk.index:
        raise LoiNghiepVu("Không tìm thấy hàng (hoặc hàng đã ngừng dùng).")
    try:
        if date.fromisoformat(hieu_luc_tu) > hom_nay():
            raise LoiNghiepVu("Ngày hiệu lực không được sau hôm nay.")
    except ValueError:
        raise LoiNghiepVu("Ngày hiệu lực không đúng.")
    h = hk.loc[ma_hang]
    cu = he_so(data, ma_hang, hieu_luc_tu)
    dvt_misa = cu["dvt_misa"] if cu is not None else ""
    if not dvt_misa:
        m = ma_misa_cua_so(data, h["so"])
        dvt_misa = m[m["ma_misa"] == h["ma_misa"]].iloc[0]["dvt_misa"]
    t = tinh_quy_doi(data, dvt_kho=h["dvt_kho"], dvt_misa=dvt_misa, cach=cach, **kw)
    if cach == "TRUC_TIEP" and not ghi_chu.strip():
        raise LoiNghiepVu("Hệ số nhập trực tiếp phải ghi nguồn (số cân / chứng từ nào) vào ô ghi chú.")
    _ghi_dong_quy_doi(store, user, ma_hang, h["dvt_kho"], dvt_misa, t, hieu_luc_tu, ghi_chu, bay_gio())
    ghi_nhat_ky(store, user["ma"], "DOI_HE_SO", f"{ma_hang} → {t['he_so']:g} từ {hieu_luc_tu}", user["nha_may"])
    return t
