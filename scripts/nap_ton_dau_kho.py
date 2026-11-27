"""Nạp TỒN ĐẦU theo số KHO (V5, cột tồn 23/11/2026) vào app, từng quy cách, đúng đơn vị kho (tấm, cây, cái…).

Mỗi (nhà máy, sổ) một phiếu kiểm kê K1 ngày 23/11/2026. Chạy lại không nạp trùng (chỉ ghi phần chênh so với tồn hiện có).
Dòng kho tồn âm: đặt 0 và liệt kê (không đoán). Hàng chưa có trong danh mục của CSDL: dừng và báo.

Dùng:
  python scripts/nap_ton_dau_kho.py danh_muc/DANH_MUC_NAP_APP_TEST.xlsx                    (theo .streamlit/secrets.toml, Google Sheets)
  python scripts/nap_ton_dau_kho.py danh_muc/DANH_MUC_NAP_APP_TEST.xlsx --local file.xlsx   (file Excel chạy thử)
  thêm --khong-ghi để chỉ xem sẽ nạp gì.
Kết quả đối chiếu: danh_muc/KIEM_TRA_TON_DAU_KHO.xlsx (tồn kho V5 so với tồn trong app, chênh phải bằng 0).
"""
import argparse
import os
import sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "scripts"))

import pandas as pd  # noqa: E402

from kho import logic as L  # noqa: E402
from scripts.cong_cu import mo  # noqa: E402

NGAY = "2026-11-23"


def nha_may_cua(sheet) -> str:
    return "NM1" if str(sheet) == "NM1" else "NM2"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("file_danh_muc")
    ap.add_argument("--local")
    ap.add_argument("--ngay", default=NGAY)
    ap.add_argument("--khong-ghi", action="store_true")
    ap.add_argument("--bao-cao", default=os.path.join(ROOT, "danh_muc", "KIEM_TRA_TON_DAU_KHO.xlsx"))
    a = ap.parse_args()

    mp = pd.read_excel(a.file_danh_muc, sheet_name="MAP_DONG_V5")
    mp["ton"] = pd.to_numeric(mp["ton_23_11"], errors="coerce").fillna(0.0)
    mp["nha_may"] = mp["sheet"].map(nha_may_cua)
    gop = mp.groupby(["nha_may", "ma_hang"], as_index=False).agg(ton=("ton", "sum"), dong=("dong", "size"),
                                                                   nguon=("sheet", lambda x: ", ".join(sorted(set(x)))))
    am = gop[gop["ton"] < 0].copy()
    gop.loc[gop["ton"] < 0, "ton"] = 0.0
    gop = gop[gop["ton"] > 0]

    store = mo(argparse.Namespace(local=a.local))
    data = store.read_all()
    hang = data["HANG_KHO"].drop_duplicates("ma_hang", keep="last").set_index("ma_hang")
    thieu = sorted(set(gop["ma_hang"]) - set(hang.index))
    if thieu:
        raise SystemExit(f"{len(thieu)} mã quy cách chưa có trong HANG_KHO của CSDL (ví dụ {thieu[:5]}). "
                         f"Hãy chạy nap_danh_muc bằng ĐÚNG file {os.path.basename(a.file_danh_muc)} trước.")
    gop["so"] = gop["ma_hang"].map(hang["so"])
    gop["ten_hang"] = gop["ma_hang"].map(hang["ten_hang"])
    gop["dvt_kho"] = gop["ma_hang"].map(hang["dvt_kho"])
    kho_ok = {(r.nha_may, r.so) for r in data["KHO_MISA"].itertuples()}
    khong_kho = gop[[(n, s) not in kho_ok for n, s in zip(gop["nha_may"], gop["so"])]]
    if not khong_kho.empty:
        print(f"CHÚ Ý: {len(khong_kho)} dòng thuộc (nhà máy, sổ) chưa khai kho MISA, bỏ qua:",
              khong_kho.groupby(["nha_may", "so"]).size().to_dict())
    gop = gop.drop(khong_kho.index)

    admin = dict(ma="NAP", ten="Nạp tồn đầu", nha_may="NM2", quyen="QUAN_TRI")
    da_ghi = []
    for nm in ("NM1", "NM2"):
        g = gop[gop["nha_may"] == nm]
        if g.empty:
            continue
        dem = pd.DataFrame({"ma_hang": g["ma_hang"], "sl_thuc_te": g["ton"]})
        phieus, _ = L.phieu_kiem_ke(data, nm, a.ngay, dem)
        for p in phieus:
            so = hang.loc[p.dong[0].ma_hang, "so"]
            p.ly_do = "Tồn đầu theo số KHO 23/11/2026 (dữ liệu thử)"
            p.ghi_chu_phieu = f"Tồn đầu theo số kho (V5 {a.ngay}) – {nm} – sổ {so}"
            p.phieu_id = f"TONDAU-KHO-{a.ngay}-{nm}-{so}"
            if a.khong_ghi:
                print(f"[không ghi] {nm} {so}: {len(p.dong)} dòng")
                continue
            sp = L.lap_phieu(store, store.read_all(), dict(admin, nha_may=nm), p)
            da_ghi.append((nm, so, len(p.dong), sp))
            print(f"Đã ghi {sp}: {len(p.dong)} dòng ({nm}, sổ {so})")
    if not da_ghi and not a.khong_ghi:
        print("Không có gì để ghi (tồn trong app đã khớp số kho).")

    # ---- Đối chiếu sau khi nạp ----
    if a.khong_ghi:
        return
    data = store.read_all()
    rows = []
    for nm in ("NM1", "NM2"):
        t = L.tinh_ton(data, nm, a.ngay).set_index("ma_hang")["ton_kho"]
        for r in gop[gop["nha_may"] == nm].itertuples():
            app = float(t.get(r.ma_hang, 0.0))
            rows.append(dict(nha_may=nm, so=r.so, ma_hang=r.ma_hang, ten_hang=r.ten_hang, dvt_kho=r.dvt_kho,
                             ton_kho_v5=r.ton, ton_app=app, chenh=round(app - r.ton, 9), nguon=r.nguon))
    dc = pd.DataFrame(rows)
    lech = dc[dc["chenh"].abs() > 1e-9]
    tom_tat = pd.DataFrame([
        ("Ngày tồn đầu", a.ngay), ("Số (hàng, nhà máy) có tồn dương", len(dc)),
        ("  NM1", int((dc["nha_may"] == "NM1").sum())), ("  NM2", int((dc["nha_may"] == "NM2").sum())),
        ("Số dòng LỆCH giữa tồn kho V5 và tồn app (phải = 0)", len(lech)),
        ("Dòng kho tồn ÂM đặt về 0 (xem sheet TON_AM_DAT_0)", len(am)),
        ("Dòng thuộc nhà máy/sổ chưa khai kho MISA (bỏ qua)", len(khong_kho))], columns=["Mục", "Giá trị"])
    with pd.ExcelWriter(a.bao_cao, engine="openpyxl") as w:
        tom_tat.to_excel(w, sheet_name="TOM_TAT", index=False)
        dc.to_excel(w, sheet_name="DOI_CHIEU", index=False)
        (am if not am.empty else pd.DataFrame(columns=am.columns)).to_excel(w, sheet_name="TON_AM_DAT_0", index=False)
        for ws in w.book.worksheets:
            ws.freeze_panes = "A2"
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = min(max(len(str(col[0].value or "")) + 2, 10), 60)
    print(f"\nĐỐI CHIẾU: {len(dc)} hàng-nhà máy, LỆCH {len(lech)} → {'KHỚP 100%' if lech.empty else 'CÓ LỆCH, xem file'}")
    print("Báo cáo:", a.bao_cao)
    if not am.empty:
        print("Tồn âm đặt 0:", am[["nha_may", "ma_hang", "ton"]].to_dict("records"))


if __name__ == "__main__":
    main()
