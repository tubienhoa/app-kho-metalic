"""Chuyển dữ liệu phát sinh từ 1 file chung (APP_KHO_PHAT_SINH_2026) sang 2 file riêng từng nhà máy.

Trước khi chạy:
  1. Tạo 2 file Google Sheets RỖNG: APP_KHO_PHAT_SINH_NM1_2026 và APP_KHO_PHAT_SINH_NM2_2026,
     chia sẻ cho tài khoản dịch vụ (quyền Người chỉnh sửa), giống file cũ.
  2. Trong .streamlit/secrets.toml, GIỮ sheet_id_phat_sinh (file cũ) và thêm:
        sheet_id_nm1 = "..."
        sheet_id_nm2 = "..."
Cách dùng (đứng ở thư mục APP_KHO):
  python scripts/chuyen_tach_nha_may.py                   # CHẠY THỬ: chỉ đếm, không ghi gì
  python scripts/chuyen_tach_nha_may.py --thuc-hien       # ghi thật vào 2 file mới
  python scripts/chuyen_tach_nha_may.py --thuc-hien --giu-tat-ca   # chuyển cả phiếu thử (mặc định chỉ chuyển tồn đầu K1)

Quy tắc an toàn: chỉ ĐỌC file cũ (không sửa, không xóa); từ chối chạy nếu file mới đã có phiếu; sau khi ghi đọc lại
và đối chiếu số dòng, tổng số lượng từng nhà máy.
Mặc định chỉ chuyển: NGUOI_DUNG, các phiếu tồn đầu K1 (loại K), KY_KHOA, TON_CHOT, HUY_PHIEU của phiếu được chuyển.
NHAT_KY cũ KHÔNG chuyển (file cũ giữ làm lưu trữ).
"""
import argparse
import os
import sys
import tomllib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from kho.schema import NHA_MAY  # noqa: E402
from kho.store import mo_store  # noqa: E402


def chuyen(cu, moi, giu_tat_ca: bool = False, thuc_hien: bool = False) -> list[str]:
    """cu: store cũ (chỉ đọc). moi: store 3 file (ghi). Trả về các dòng báo cáo; ném lỗi nếu đối chiếu sai."""
    bao = []
    if not getattr(moi, "tach_nha_may", False):
        raise SystemExit("Store mới phải là cấu hình 3 file (có sheet_id_nm1, sheet_id_nm2).")
    d_cu = cu.read_all()
    gd = d_cu["GIAO_DICH"]
    if thuc_hien:
        moi.khoi_tao()    # tạo sheet còn thiếu ở 3 file (file danh mục: chỉ thêm HANG_NHA_MAY nếu chưa có)
    try:
        d_moi = moi.read_all()
    except Exception:
        if thuc_hien:
            raise
        d_moi = None   # chạy thử khi file mới còn trống, chưa có sheet: coi như rỗng
        bao.append("File nhà máy mới chưa có các sheet (sẽ được tạo khi chạy với --thuc-hien).")
    if d_moi is not None and (len(d_moi["GIAO_DICH"]) or len(d_moi["NGUOI_DUNG"])):
        raise SystemExit("File nhà máy mới đã có dữ liệu (GIAO_DICH/NGUOI_DUNG). Từ chối chạy để không ghi trùng.")

    chon = gd if giu_tat_ca else gd[gd["loai"] == "K1"]
    pid = set(chon["phieu_id"])
    huy = d_cu["HUY_PHIEU"]
    huy = huy[huy["phieu_id"].isin(pid)]
    bang = {"GIAO_DICH": chon, "HUY_PHIEU": huy, "NGUOI_DUNG": d_cu["NGUOI_DUNG"],
            "KY_KHOA": d_cu["KY_KHOA"], "TON_CHOT": d_cu["TON_CHOT"]}
    bao.append(f"File cũ: {len(gd)} dòng giao dịch ({gd['phieu_id'].nunique()} phiếu). Sẽ chuyển "
               f"{len(chon)} dòng ({chon['phieu_id'].nunique()} phiếu"
               f"{'' if giu_tat_ca else ', chỉ tồn đầu K1'}).")
    ghi = {}
    for ten, df in bang.items():
        for nm in NHA_MAY:
            if ten == "HUY_PHIEU":
                con = df[df["so_phieu"].str[:3] == nm]
            else:
                con = df[df["nha_may"] == nm]
            ghi[(ten, nm)] = con
            bao.append(f"  {ten:<11} → {nm}: {len(con)} dòng")
        thieu = len(df) - sum(len(ghi[(ten, nm)]) for nm in NHA_MAY)
        if thieu:
            raise SystemExit(f"{ten}: {thieu} dòng không thuộc NM1 hoặc NM2, dừng lại để kiểm tra.")
    if not thuc_hien:
        bao.append("CHẠY THỬ – chưa ghi gì. Thêm --thuc-hien để ghi thật.")
        return bao
    for (ten, nm), df in ghi.items():
        rows = df.to_dict("records")
        for r in rows:
            r["nha_may"] = nm     # để store định tuyến (HUY_PHIEU không có cột này)
        moi.append(ten, rows)
    # đối chiếu: đọc lại từ file mới
    d2 = moi.read_all()
    for ten, df in bang.items():
        assert len(d2[ten]) == len(df), f"{ten}: đọc lại {len(d2[ten])} dòng, mong đợi {len(df)}"
    for nm in NHA_MAY:
        a = chon[chon["nha_may"] == nm]
        b = d2["GIAO_DICH"][d2["GIAO_DICH"]["nha_may"] == nm]
        assert abs((a["dau"] * a["sl_kho"]).sum() - (b["dau"] * b["sl_kho"]).sum()) < 1e-6, f"{nm}: tổng số lượng lệch"
    bao.append("ĐÃ GHI và đối chiếu ĐẠT (số dòng từng sheet, tổng số lượng từng nhà máy). File cũ giữ nguyên làm lưu trữ.")
    return bao


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--thuc-hien", action="store_true")
    ap.add_argument("--giu-tat-ca", action="store_true")
    a = ap.parse_args()
    with open(os.path.join(ROOT, ".streamlit", "secrets.toml"), "rb") as f:
        cfg = dict(tomllib.load(f)["luu_tru"])
    if not (cfg.get("sheet_id_nm1") and cfg.get("sheet_id_nm2") and cfg.get("sheet_id_phat_sinh")):
        raise SystemExit("secrets.toml cần có đủ sheet_id_phat_sinh (file cũ), sheet_id_nm1, sheet_id_nm2.")
    cu_cfg = {k: v for k, v in cfg.items() if k not in ("sheet_id_nm1", "sheet_id_nm2")}
    for dong in chuyen(mo_store(cu_cfg), mo_store(cfg), a.giu_tat_ca, a.thuc_hien):
        print(dong)


if __name__ == "__main__":
    main()
