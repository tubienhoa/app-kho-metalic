"""Tạo file dữ liệu CHẠY THỬ (dữ liệu giả, không phải số thật) để thử app trên máy.

Chạy:  python scripts/tao_du_lieu_thu.py
Tạo:   du_lieu_thu/APP_KHO_DATA_THU.xlsx  (xóa file cũ nếu có)
Người dùng thử: QT (quản trị) PIN 1234; K2 (kho NM2) PIN 1111; K1 (kho NM1) PIN 2222
"""
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from kho.auth import hash_pin  # noqa: E402
from kho.store import LocalStore  # noqa: E402

PATH = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                    "du_lieu_thu", "APP_KHO_DATA_THU.xlsx")


def tao(path=PATH):
    if os.path.exists(path):
        os.remove(path)
    s = LocalStore(path)
    t = "2026-11-27 08:00:00"
    s.append("NGUOI_DUNG", [
        dict(ma="QT", ten="Quản trị (thử)", nha_may="NM2", quyen="QUAN_TRI", pin_hash=hash_pin("1234"), hoat_dong="1", cap_nhat_luc=t),
        dict(ma="K2", ten="Kho NM2 (thử)", nha_may="NM2", quyen="KHO", pin_hash=hash_pin("1111"), hoat_dong="1", cap_nhat_luc=t),
        dict(ma="K1", ten="Kho NM1 (thử)", nha_may="NM1", quyen="KHO", pin_hash=hash_pin("2222"), hoat_dong="1", cap_nhat_luc=t),
    ])
    s.append("KHO_MISA", [dict(nha_may="NM2", so="MTL", kho_misa="152.SX"),
                          dict(nha_may="NM1", so="MTL", kho_misa="152.GC"),
                          dict(nha_may="NM1", so="CMA", kho_misa="CMA-THU")])
    s.append("MA_MISA", [dict(so="MTL", ma_misa="THU_THEPTAM3", ten_misa="Thép tấm 3 ly (thử)", dvt_misa="kg"),
                         dict(so="MTL", ma_misa="THU_KHIOXY", ten_misa="Khí oxy (thử)", dvt_misa="bình"),
                         dict(so="CMA", ma_misa="THU_THEPTAM3", ten_misa="Thép tấm 3 ly CMA (thử)", dvt_misa="kg")])
    s.append("HANG_KHO", [
        dict(ma_hang="TT3-1220x2500", ten_hang="Thép tấm 3 x 1220 x 2500mm (thử)", nhom="Thép tấm", dvt_kho="tấm",
             so="MTL", ma_misa="THU_THEPTAM3", vat_lieu="Thép", quy_cach="3x1220x2500", hoat_dong="1"),
        dict(ma_hang="TT3-1220x600", ten_hang="Thép tấm 3 x 1220 x 600mm (thử, sau cắt)", nhom="Thép tấm", dvt_kho="tấm",
             so="MTL", ma_misa="THU_THEPTAM3", vat_lieu="Thép", quy_cach="3x1220x600", hoat_dong="1"),
        dict(ma_hang="OXY", ten_hang="Khí oxy (thử)", nhom="Khí", dvt_kho="bình",
             so="MTL", ma_misa="THU_KHIOXY", hoat_dong="1"),
        dict(ma_hang="CMA-TT3-1220x2500", ten_hang="Thép tấm 3 x 1220 x 2500mm CMA (thử)", nhom="Thép tấm", dvt_kho="tấm",
             so="CMA", ma_misa="THU_THEPTAM3", vat_lieu="Thép", quy_cach="3x1220x2500", hoat_dong="1"),
        dict(ma_hang="CHUA-HS", ten_hang="Hàng chưa duyệt hệ số (thử)", nhom="Khác", dvt_kho="cái",
             so="MTL", ma_misa="THU_KHIOXY", hoat_dong="1"),
    ])
    # 3 x 1.22 x 2.5 x 7.85 = 71.8275 kg/tấm ; 3 x 1.22 x 0.6 x 7.85 = 17.2386
    s.append("QUY_DOI", [
        dict(ma_hang="TT3-1220x2500", dvt_kho="tấm", dvt_misa="kg", he_so=71.8275, cach_tinh="KLR thép 7,85",
             nguoi_duyet="QT", ngay_duyet="2026-11-27", hieu_luc_tu=""),
        dict(ma_hang="TT3-1220x600", dvt_kho="tấm", dvt_misa="kg", he_so=17.2386, cach_tinh="KLR thép 7,85",
             nguoi_duyet="QT", ngay_duyet="2026-11-27", hieu_luc_tu=""),
        dict(ma_hang="OXY", dvt_kho="bình", dvt_misa="bình", he_so=1, cach_tinh="Cùng ĐVT",
             nguoi_duyet="QT", ngay_duyet="2026-11-27", hieu_luc_tu=""),
        dict(ma_hang="CMA-TT3-1220x2500", dvt_kho="tấm", dvt_misa="kg", he_so=71.8275, cach_tinh="KLR thép 7,85",
             nguoi_duyet="QT", ngay_duyet="2026-11-27", hieu_luc_tu=""),
        dict(ma_hang="CHUA-HS", dvt_kho="cái", dvt_misa="bình", he_so=2, cach_tinh="Đề xuất", nguoi_duyet="", hieu_luc_tu=""),
    ])
    s.append("DOI_TAC", [dict(ma_doi_tac="NCC01", ten_doi_tac="NCC thép (thử)", loai="NCC"),
                         dict(ma_doi_tac="GC01", ten_doi_tac="Xưởng mạ (thử)", loai="GIA_CONG"),
                         dict(ma_doi_tac="PL01", ten_doi_tac="Vựa phế liệu (thử)", loai="MUA_PHE_LIEU")])
    s.append("BO_PHAN", [dict(nha_may="NM2", ten_bo_phan="Tổ chấn"), dict(nha_may="NM2", ten_bo_phan="Tổ hàn"),
                         dict(nha_may="NM2", ten_bo_phan="Tổ cắt laser"), dict(nha_may="NM1", ten_bo_phan="Tổ hàn")])
    return s


if __name__ == "__main__":
    tao()
    print("Đã tạo", PATH)
