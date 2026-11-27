"""Dựng FILE DANH MỤC NẠP APP từ V5 + phiếu kiểm v3 (CHỈ ĐỌC, không sửa 2 file nguồn; ghi SHA256 trước/sau).

Quyết định của người dùng (27/11/2026) được áp dụng:
  - Chỉ nạp dòng kho ĐÃ CÓ MÃ MISA và CÓ HỆ SỐ QUY ĐỔI rõ ràng (cùng ĐVT, theo kích thước/khối lượng riêng,
    theo số ghi trong tên, hoặc quy tắc đã duyệt 25/11 QT03/QT04/đồng nghĩa).
  - LOẠI RA: 73 dòng cần cân thực tế (chờ kiểm kê), 511 dòng chưa có mã, và các dòng không đủ dữ kiện
    (hệ số chưa có, mã chưa có ĐVT MISA, ô ghi nhiều mã…). Tất cả liệt kê ở sheet CHUA_NAP kèm lý do.
  - Hàng tồn 0 đã có mã: vẫn nạp.

Thêm --ma-tam (chỉ cho file TEST): nạp luôn các dòng kho CÒN TỒN chưa vào được danh mục chính (chưa có mã / thiếu
hệ số) dưới mã tạm "CHUA_CO_MA" (đơn vị theo kho, hệ số 1), đánh dấu rõ là TẠM, để kho thử đủ vật tư.

Chạy:  python scripts/dung_danh_muc.py            -> danh_muc/DANH_MUC_NAP_APP.xlsx
       python scripts/dung_danh_muc.py --ma-tam   -> danh_muc/DANH_MUC_NAP_APP_TEST.xlsx
Rồi:   python scripts/cong_cu.py nap_danh_muc danh_muc/DANH_MUC_NAP_APP.xlsx   (nạp vào Google Sheets)
"""
import argparse
import hashlib
import os
import re
import sys
import unicodedata
from datetime import datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import openpyxl  # noqa: E402
import pandas as pd  # noqa: E402

from kho import quydoi as Q  # noqa: E402
from kho.schema import DUNG_SAI, SHEETS  # noqa: E402

V5 = r"D:\CLAUDE_DATA_TONKHO\T2_23112026\Danh muc du lieu_V5.xlsx"
V3 = r"D:\CLAUDE_DATA_TONKHO\T2_23112026\So ton kho_23112026_v3.xlsx"
NGAY_DUYET = "2026-11-25"
NGUOI_DUYET = "QT 25/11/2026"


def sha(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for b in iter(lambda: f.read(1 << 20), b""):
            h.update(b)
    return h.hexdigest()


def nfc(s):
    return unicodedata.normalize("NFC", str(s)).strip() if s is not None else ""


def norm_ma(v):
    s = str(v).strip().upper()
    s = re.sub(r"\s*-\s*", "-", s)
    return re.sub(r"\s+", "", s)


def norm_ten(v):
    return re.sub(r"\s+", " ", nfc(v).lower())


def doc_misa(v5):
    """(sổ, mã chuẩn hóa) -> dict(ma, ten, dvt). Mã lặp ở nhiều kho: lấy DÒNG CUỐI (theo ý người dùng 'thứ tự lớn nhất')."""
    wb = openpyxl.load_workbook(v5, data_only=True, read_only=True)
    out = {}
    for i, r in enumerate(wb["TỔNG HỢP TỒN KHO MTL"].iter_rows(min_row=6, values_only=True), 6):
        if r[2]:
            out[("MTL", norm_ma(r[2]))] = dict(ma=nfc(r[2]), ten=nfc(r[3]), dvt=nfc(r[5]))
    for i, r in enumerate(wb["TỔNG HỢP TỒN KHO CMA"].iter_rows(min_row=6, values_only=True), 6):
        if r[1]:
            out[("CMA", norm_ma(r[1]))] = dict(ma=nfc(r[1]), ten=nfc(r[2]), dvt=nfc(r[3]))
    wb.close()
    return out


NGAY_MISA = "2026-11-23"
MA_TAM = "CHUA_CO_MA"


def doc_ton_misa(v5):
    """Số tồn cuối kỳ MISA (SL, giá trị) theo (sổ, kho, mã) của 3 kho trong phạm vi: 152.SX, 152.GC (MTL) và CMA."""
    wb = openpyxl.load_workbook(v5, data_only=True, read_only=True)
    out = []
    num = lambda v: float(v) if isinstance(v, (int, float)) else ""
    for r in wb["TỔNG HỢP TỒN KHO MTL"].iter_rows(min_row=6, values_only=True):
        if r[2] and str(r[1]).strip() in ("152.SX", "152.GC"):
            out.append(dict(ngay=NGAY_MISA, so="MTL", kho_misa=str(r[1]).strip(), ma_misa=nfc(r[2]), sl_misa=num(r[12]),
                            dvt_misa=nfc(r[5]), gia_tri=num(r[13])))
    for r in wb["TỔNG HỢP TỒN KHO CMA"].iter_rows(min_row=6, values_only=True):
        if r[1]:
            out.append(dict(ngay=NGAY_MISA, so="CMA", kho_misa="CMA", ma_misa=nfc(r[1]), sl_misa=num(r[10]),
                            dvt_misa=nfc(r[3]), gia_tri=num(r[11])))
    wb.close()
    return out


def ly_do_loai(cach: str) -> str:
    c = str(cach)
    if c.startswith("Không có số kg/"):
        return "Cần cân thực tế (chưa có kg/ĐVT) – chờ kiểm kê"
    if "khác ĐVT MISA" in c:
        return "ĐVT kho khác ĐVT MISA, chưa có hệ số"
    if c.startswith("Mã không có trong báo cáo MISA"):
        return "Mã chưa có trong báo cáo MISA (không biết ĐVT MISA)"
    if "lệch nhau quá 2%" in c:
        return "Các nguồn trọng lượng lệch nhau quá 2%"
    if c.startswith("Kho ghi kg, MISA ghi"):
        return "Kho ghi kg, MISA ghi đơn vị khác, chưa có hệ số"
    return "Chưa có hệ số: " + c[:80]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--v5", default=V5)
    ap.add_argument("--v3", default=V3)
    ap.add_argument("--out", default=None)
    ap.add_argument("--ma-tam", action="store_true", help="nạp dòng còn tồn chưa có mã/hệ số dưới mã tạm (chỉ file TEST)")
    a = ap.parse_args()
    if a.out is None:
        a.out = os.path.join(ROOT, "danh_muc", "DANH_MUC_NAP_APP_TEST.xlsx" if a.ma_tam else "DANH_MUC_NAP_APP.xlsx")
    h5a, h3a = sha(a.v5), sha(a.v3)

    misa = doc_misa(a.v5)
    ct = pd.read_excel(a.v3, sheet_name="CHI_TIET_KHO")
    chua_ma = pd.read_excel(a.v3, sheet_name="CHUA_CO_MA")
    bang_nhom = pd.read_excel(a.v3, sheet_name="BANG_NHOM_HANG")
    tien_to = sorted(((str(r["Tiền tố mã kế toán"]).strip().upper(), str(r["Nhóm hàng"]).strip())
                      for _, r in bang_nhom.iterrows()), key=lambda x: -len(x[0]))

    def nhom_cua(ma):
        for t, n in tien_to:
            if ma.upper().startswith(t):
                return n
        return "Khác"

    hs_col = "Hệ số (ĐVT MISA / 1 ĐVT kho)"
    sanSang, chuaNap = [], []
    for _, r in ct.iterrows():
        so = str(r["Sổ"]).strip()
        o_ma = str(r["Ô mã kế toán"]).strip()
        base = dict(sheet=r["Sheet (V5)"], dong=int(r["Dòng (bấm mở V5)"]), o_ma=o_ma, ten=nfc(r["Tên hàng (kho)"]),
                    dvt_kho=nfc(r["ĐVT kho"]), ton=r["Tồn kho 23/11 (ĐVT kho)"], so=so, cach=str(r["Cách quy đổi"]),
                    trang_thai=str(r["Trạng thái map (V5)"]))
        if len([x for x in re.split(r"[;,\n]+", o_ma) if x.strip()]) > 1:   # như v3: chỉ ; , xuống dòng tách mã
            chuaNap.append(dict(base, ly_do="Ô mã ghi nhiều mã (không biết chọn mã nào)"))
            continue
        m = misa.get((so, norm_ma(o_ma)))
        hs = pd.to_numeric(r[hs_col], errors="coerce")
        if not (hs > 0):
            chuaNap.append(dict(base, ly_do=ly_do_loai(r["Cách quy đổi"])))
            continue
        if m is None:
            chuaNap.append(dict(base, ly_do="Mã không có trong danh sách mã MISA của sổ này"))
            continue
        sanSang.append(dict(base, he_so=float(hs), ma_misa=m["ma"], ten_misa=m["ten"], dvt_misa=m["dvt"]))

    # ---- Gộp dòng kho trùng quy cách thành 1 hàng ----
    nhom_gop = {}
    for d in sanSang:
        k = (d["so"], d["ma_misa"], norm_ten(d["ten"]), d["dvt_kho"].lower())
        nhom_gop.setdefault(k, []).append(d)
    cac_hang, canh_bao, map_dong = [], [], []
    for stt, (k, ds) in enumerate(sorted(nhom_gop.items(), key=lambda x: (x[0][0], x[0][1], x[0][2])), 1):
        d = ds[0]
        he_khac = {round(x["he_so"], 9) for x in ds}
        if len(he_khac) > 1:  # không đoán: các dòng cùng quy cách mà hệ số khác nhau -> loại, báo
            for x in ds:
                chuaNap.append(dict(x, ly_do=f"Các dòng cùng quy cách có hệ số khác nhau {sorted(he_khac)}"))
            continue
        ma_hang = f"Q{len(cac_hang) + 1:05d}"
        dvt_kho = d["dvt_kho"].strip().lower()
        kg_ok = d["dvt_misa"].strip().lower() == "kg"
        # Hàng tính bằng kg (ĐVT kho = kg) có hệ số 1: không có "1 tấm/1 cây" để so với kg lý thuyết.
        ctt = Q.doc_kich_thuoc(d["ten"]) if kg_ok and dvt_kho != "kg" else None
        vl, klr = Q.vat_lieu(d["ten"]) if ctt else (None, None)
        lt, kt = "", [""] * 4
        if ctt and klr:
            lt = round(Q.kg_ly_thuyet(ctt["hinh_dang"], klr, ctt["kt"]), 6)
            kt = list(ctt["kt"]) + [""] * (4 - len(ctt["kt"]))
            if Q.vuot_dung_sai(d["he_so"], lt):
                canh_bao.append(dict(ma_hang=ma_hang, ten=d["ten"], dvt_kho=dvt_kho, he_so=d["he_so"], kg_ly_thuyet=lt,
                                     lech_pct=round(Q.lech(d["he_so"], lt) * 100, 2), cach=d["cach"],
                                     ghi_chu=f"Hệ số lệch kg lý thuyết quá ±{DUNG_SAI * 100:g}% (nạp theo hệ số v3, cần xem lại)"))
        nguon = "; ".join(f"{x['sheet']}!{x['dong']}" for x in ds)
        ghi_chu = "kế toán xem quy cách" if any("xem quy cách" in x["trang_thai"] for x in ds) else ""
        cac_hang.append(dict(
            hang=dict(ma_hang=ma_hang, ten_hang=d["ten"], nhom=nhom_cua(d["ma_misa"]), dvt_kho=dvt_kho, so=d["so"],
                      ma_misa=d["ma_misa"], vat_lieu=vl or "", quy_cach="x".join(f"{x:g}" for x in ctt["kt"]) if ctt else "",
                      hoat_dong="1", ghi_chu=ghi_chu, nguon=nguon),
            qd=dict(ma_hang=ma_hang, dvt_kho=dvt_kho, dvt_misa=d["dvt_misa"], he_so=d["he_so"], cach_tinh=d["cach"],
                    nguoi_duyet=NGUOI_DUYET, ngay_duyet=NGAY_DUYET, hieu_luc_tu="",
                    ghi_chu="Nạp từ phiếu kiểm v3 (V5 23/11); duyệt theo quy tắc QT01/QT03/QT04/đồng nghĩa 25/11",
                    hinh_dang=ctt["hinh_dang"] if ctt else "", vat_lieu=vl or "", klr=klr or "", kt_a=kt[0], kt_b=kt[1],
                    kt_c=kt[2], kt_d=kt[3], kg_ly_thuyet=lt, kg_can="")))
        for x in ds:
            map_dong.append(dict(sheet=x["sheet"], dong=x["dong"], o_ma=x["o_ma"], ten=x["ten"], ma_hang=ma_hang,
                                 ton_23_11=x["ton"], loai="chính thức"))

    # ---- (chỉ file TEST) nạp dòng còn tồn chưa vào được dưới mã tạm ----
    tam_hang, tam_so_dong, tam_bo_qua = [], 0, []
    if a.ma_tam:
        nguon_tam = []
        for x in chuaNap:
            nguon_tam.append(dict(sheet=x["sheet"], dong=x["dong"], o_ma=x["o_ma"], ten=x["ten"], dvt=x["dvt_kho"],
                                  ton=x["ton"], so=x["so"], ly_do=x["ly_do"]))
        for _, r in chua_ma.iterrows():
            so_dk = str(r.get("Sổ / kho MISA dự kiến", "")).split()
            nguon_tam.append(dict(sheet=r["Sheet (V5)"], dong=int(r["Dòng (bấm mở V5)"]), o_ma=str(r.get("Ô mã kế toán (V5)", "")),
                                  ten=nfc(r["Tên hàng (kho)"]) if pd.notna(r["Tên hàng (kho)"]) else "",
                                  dvt=nfc(r["ĐVT kho"]) if pd.notna(r["ĐVT kho"]) else "", ton=r["Tồn kho 23/11"],
                                  so="CMA" if so_dk[:1] == ["CMA"] else "MTL", ly_do="Chưa có mã MISA"))
        gop = {}
        for x in nguon_tam:
            ton = pd.to_numeric(x["ton"], errors="coerce")
            if not (ton > 0):
                continue          # chỉ nạp dòng còn tồn dương (tồn âm / trống: không nạp, không đoán)
            ten = x["ten"] or f"(chưa có tên – V5 {x['sheet']} dòng {x['dong']})"
            dvt = (x["dvt"] or "").strip().lower()
            if not dvt:
                tam_bo_qua.append(dict(x, ly_do="Dòng kho không ghi ĐVT – không nạp được"))
                continue
            gop.setdefault((x["so"], norm_ten(ten), dvt), []).append(dict(x, ten=ten, dvt=dvt, ton=float(ton)))
        for k, ds in sorted(gop.items(), key=lambda kv: (kv[0][0], kv[0][1])):
            d = ds[0]
            ma_hang = f"Q{len(cac_hang) + 1:05d}"
            ma_v5 = "; ".join(sorted({x["o_ma"] for x in ds if x["o_ma"] and x["o_ma"].lower() != "nan"}))
            ly = "; ".join(sorted({x["ly_do"] for x in ds}))
            ghi = f"TẠM (dữ liệu thử): {ly}." + (f" Mã ghi trong V5: {ma_v5}." if ma_v5 else "")
            cac_hang.append(dict(
                hang=dict(ma_hang=ma_hang, ten_hang=d["ten"], nhom="CHƯA CÓ MÃ (TẠM)", dvt_kho=d["dvt"], so=d["so"], ma_misa=MA_TAM,
                          vat_lieu="", quy_cach="", hoat_dong="1", ghi_chu=ghi,
                          nguon="; ".join(f"{x['sheet']}!{x['dong']}" for x in ds)),
                qd=dict(ma_hang=ma_hang, dvt_kho=d["dvt"], dvt_misa=d["dvt"], he_so=1.0,
                        cach_tinh="TẠM: hệ số 1 theo ĐVT kho (chưa có mã/hệ số thật)", nguoi_duyet="TAM_THU",
                        ngay_duyet=NGAY_DUYET, hieu_luc_tu="", ghi_chu=ghi, hinh_dang="", vat_lieu="", klr="", kt_a="", kt_b="",
                        kt_c="", kt_d="", kg_ly_thuyet="", kg_can="")))
            tam_hang.append(ma_hang)
            for x in ds:
                tam_so_dong += 1
                map_dong.append(dict(sheet=x["sheet"], dong=x["dong"], o_ma=x["o_ma"], ten=x["ten"], ma_hang=ma_hang,
                                     ton_23_11=x["ton"], loai="TẠM"))

    # ---- Các sheet danh mục ----
    ma_misa_rows = [dict(so=so, ma_misa=m["ma"], ten_misa=m["ten"], dvt_misa=m["dvt"]) for (so, _), m in misa.items()]
    if tam_hang:
        for so_t in sorted({c["hang"]["so"] for c in cac_hang if c["hang"]["ma_misa"] == MA_TAM}):
            ma_misa_rows.append(dict(so=so_t, ma_misa=MA_TAM, ten_misa="Hàng chưa có mã MISA (mã tạm, dữ liệu thử)", dvt_misa=""))
    kho_misa = [dict(nha_may="NM2", so="MTL", kho_misa="152.SX"), dict(nha_may="NM1", so="MTL", kho_misa="152.GC"),
                dict(nha_may="NM1", so="CMA", kho_misa="CMA")]
    klr_rows = [dict(vat_lieu=k, klr=v, ghi_chu="QT01 (duyệt 25/11/2026)", nguoi="QT01", thoi_diem=NGAY_DUYET + " 00:00:00")
                for k, v in Q.KLR_MAC_DINH.items()]
    dong_nghia = [dict(dvt_goc=g, dvt_chuan=c, ghi_chu=n) for g, c, n in [
        ("pcs", "cái", "Đồng nghĩa 1:1 (duyệt 25/11)"), ("pcs", "con", "Đồng nghĩa 1:1 (duyệt 25/11)"),
        ("cái", "chiếc", "Đồng nghĩa 1:1 (duyệt 25/11)"), ("cái", "con", "QT04: 1 Cái = 1 Con"),
        ("cái", "thùng", "QT03: 1 Cái (thùng carton) = 1 Thùng"), ("set", "bộ", "Đồng nghĩa 1:1 (duyệt 25/11)"),
        ("sét", "bộ", "Đồng nghĩa 1:1 (duyệt 25/11)"), ("bộ/cặp", "cặp", "Đồng nghĩa 1:1 (duyệt 25/11)"),
        ("cuộn nhỏ", "cuộn", "Đồng nghĩa 1:1 (duyệt 25/11)"), ("thanh", "cái", "Đồng nghĩa 1:1 (duyệt 25/11)")]]
    bo_phan = [dict(nha_may="NM2", ten_bo_phan=t) for t in
               ["CNC phay tiện", "Đóng gói", "HVS", "Bảo trì", "CNC lazer", "BHLĐ", "Chung", "Sơn", "Chấn", "KTDA", "QC", "Kho"]]

    cot = lambda n: SHEETS[n]
    df = lambda rows, n: pd.DataFrame(rows, columns=cot(n))
    dm = {"KHO_MISA": df(kho_misa, "KHO_MISA"), "MA_MISA": df(ma_misa_rows, "MA_MISA"),
          "HANG_KHO": df([x["hang"] for x in cac_hang], "HANG_KHO"), "QUY_DOI": df([x["qd"] for x in cac_hang], "QUY_DOI"),
          "KHOI_LUONG_RIENG": df(klr_rows, "KHOI_LUONG_RIENG"), "DONG_NGHIA_DVT": df(dong_nghia, "DONG_NGHIA_DVT"),
          "DOI_TAC": df([], "DOI_TAC"), "BO_PHAN": df(bo_phan, "BO_PHAN"),
          "TON_MISA_THAM_KHAO": df(doc_ton_misa(a.v5), "TON_MISA_THAM_KHAO")}

    # ---- Báo cáo kiểm tra ----
    cn = pd.DataFrame(chuaNap)
    cn_cot = ["sheet", "dong", "o_ma", "ten", "dvt_kho", "ton", "so", "ly_do", "cach"]
    cn = cn[cn_cot] if not cn.empty else pd.DataFrame(columns=cn_cot)
    cn = cn.rename(columns={"sheet": "Sheet V5", "dong": "Dòng V5", "o_ma": "Ô mã kế toán", "ten": "Tên hàng (kho)",
                            "dvt_kho": "ĐVT kho", "ton": "Tồn 23/11", "so": "Sổ", "ly_do": "Lý do chưa nạp",
                            "cach": "Ghi chú quy đổi (v3)"})
    cm = chua_ma.rename(columns={"Ô mã kế toán (V5)": "Ô mã kế toán", "Tồn kho 23/11": "Tồn 23/11"}).assign(
        **{"Lý do chưa nạp": "Chưa có mã MISA – chờ file dữ liệu chuẩn người dùng sẽ báo"})
    cm = cm[["Sheet (V5)", "Dòng (bấm mở V5)", "Ô mã kế toán", "Tên hàng (kho)", "ĐVT kho", "Tồn 23/11", "Lý do chưa nạp",
             "Trạng thái map (V5)"]].rename(columns={"Sheet (V5)": "Sheet V5", "Dòng (bấm mở V5)": "Dòng V5"})
    chua_nap_all = pd.concat([cn, cm], ignore_index=True)
    if a.ma_tam:   # trong file TEST: dòng đã nạp dưới mã tạm ghi rõ, dòng bị loại hẳn ghi lý do
        chua_nap_all["Đã nạp mã tạm trong file TEST?"] = "Có (nếu còn tồn > 0)"
        if tam_bo_qua:
            chua_nap_all = pd.concat([chua_nap_all, pd.DataFrame([{"Sheet V5": x["sheet"], "Dòng V5": x["dong"], "Ô mã kế toán": x["o_ma"],
                "Tên hàng (kho)": x["ten"], "ĐVT kho": x["dvt"], "Tồn 23/11": x["ton"], "Lý do chưa nạp": x["ly_do"]} for x in tam_bo_qua])], ignore_index=True)
    ly_do_dem = chua_nap_all["Lý do chưa nạp"].value_counts()

    nhom_dem = dm["HANG_KHO"]["nhom"].value_counts()
    shape_dem = dm["QUY_DOI"]["hinh_dang"].replace("", "(không có kích thước)").value_counts()
    cach_dem = dm["QUY_DOI"]["cach_tinh"].map(lambda s: "Cùng ĐVT" if s == "Cùng ĐVT" else
                                              ("Theo kg (khối lượng riêng / tên hàng)" if "kg" in s else "Quy tắc đã duyệt 25/11")).value_counts()
    tom_tat = [("DANH MỤC NẠP APP – dựng " + datetime.now().strftime("%d/%m/%Y %H:%M"), ""),
               ("Nguồn V5", a.v5), ("SHA256 V5 trước / sau", f"{h5a[:16]}… / (kiểm sau khi ghi file)"),
               ("Nguồn phiếu kiểm v3", a.v3), ("SHA256 v3 trước / sau", f"{h3a[:16]}… / (kiểm sau khi ghi file)"), ("", ""),
               ("Dòng kho đã có mã trong v3 (CHI_TIET_KHO)", len(ct)),
               ("  Có hệ số rõ ràng → nạp", len(sanSang)),
               ("  Gộp các dòng trùng quy cách → số HÀNG nạp vào HANG_KHO", len(cac_hang)),
               ("  Không nạp (đủ dữ kiện không) – xem CHUA_NAP", len(chuaNap)),
               ("Dòng kho còn tồn nhưng chưa có mã (v3 CHUA_CO_MA)" + (" – nạp dưới mã tạm (file TEST)" if a.ma_tam else " – không nạp"), len(chua_ma)),
               ("Dòng nạp dưới mã tạm CHUA_CO_MA (chỉ file TEST)", tam_so_dong),
               ("  → số hàng tạm", len(tam_hang)), ("  Dòng còn tồn không nạp được (không ghi ĐVT)", len(tam_bo_qua)), ("", ""),
               ("Mã MISA đưa vào MA_MISA (MTL + CMA)", len(ma_misa_rows)),
               ("Hàng có đọc được kích thước → tự tính kg lý thuyết", int((dm["QUY_DOI"]["hinh_dang"] != "").sum())),
               ("Cảnh báo hệ số lệch kg lý thuyết > ±2% (sheet CANH_BAO)", len(canh_bao)), ("", ""), ("LÝ DO CHƯA NẠP", "số dòng")]
    tom_tat += [(k, int(v)) for k, v in ly_do_dem.items()] + [("", ""), ("HÀNG NẠP THEO HÌNH DẠNG", "")]
    tom_tat += [(k, int(v)) for k, v in shape_dem.items()] + [("", ""), ("HÀNG NẠP THEO CÁCH QUY ĐỔI", "")]
    tom_tat += [(k, int(v)) for k, v in cach_dem.items()] + [("", ""), ("HÀNG NẠP THEO NHÓM", "")]
    tom_tat += [(k, int(v)) for k, v in nhom_dem.items()]
    tom_tat += [("", ""), ("Số MISA tham khảo (TON_MISA_THAM_KHAO) ngày " + NGAY_MISA, len(dm["TON_MISA_THAM_KHAO"]))]
    rv = {"TOM_TAT": pd.DataFrame(tom_tat, columns=["Mục", "Giá trị"]), "CHUA_NAP": chua_nap_all,
          "CANH_BAO": pd.DataFrame(canh_bao), "MAP_DONG_V5": pd.DataFrame(map_dong)}

    os.makedirs(os.path.dirname(a.out), exist_ok=True)
    with pd.ExcelWriter(a.out, engine="openpyxl") as w:
        for n in ("TOM_TAT", "CHUA_NAP", "CANH_BAO", "MAP_DONG_V5"):
            rv[n].to_excel(w, sheet_name=n, index=False)
        for n, d in dm.items():
            d.to_excel(w, sheet_name=n, index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            for row in ws.iter_rows():
                for c in row:
                    if isinstance(c.value, str) and c.value.startswith("="):
                        c.data_type = "s"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(col[0].value or "")) + 2, 10), 60)
    h5b, h3b = sha(a.v5), sha(a.v3)
    assert h5a == h5b and h3a == h3b, "NGUỒN BỊ THAY ĐỔI TRONG LÚC CHẠY!"
    wb = openpyxl.load_workbook(a.out)
    ws = wb["TOM_TAT"]
    for row in ws.iter_rows(min_row=2):
        if row[0].value == "SHA256 V5 trước / sau":
            row[1].value = f"{h5a} → {h5b} (không đổi)"
        if row[0].value == "SHA256 v3 trước / sau":
            row[1].value = f"{h3a} → {h3b} (không đổi)"
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 90
    wb.save(a.out)
    print(f"Đã dựng {a.out}")
    for k, v in tom_tat[6:17]:
        print(f"  {k}: {v}")
    print("SHA256 V5/v3 trước = sau:", h5a == h5b, h3a == h3b)


if __name__ == "__main__":
    main()
