"""Công cụ dòng lệnh cho quản trị.

  python scripts/cong_cu.py khoi_tao              # tạo các sheet + tiêu đề trong Google Sheets (theo .streamlit/secrets.toml)
  python scripts/cong_cu.py mau_danh_muc FILE.xlsx # tạo file Excel mẫu để điền danh mục
  python scripts/cong_cu.py nap_danh_muc FILE.xlsx # nạp danh mục từ file Excel (chỉ nạp vào sheet đang TRỐNG)
  python scripts/cong_cu.py sao_luu THU_MUC        # xuất mọi sheet ra CSV (dùng cho sao lưu)
  python scripts/cong_cu.py kiem_tra_ket_noi       # thử kết nối thật: đủ sheet, ghi 1 dòng thử vào NHAT_KY và đọc lại

Thêm --local DUONG_DAN.xlsx để làm với file chạy thử thay cho Google Sheets.
"""
import argparse
import os
import sys
import tomllib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

import pandas as pd  # noqa: E402

from kho.schema import NUMERIC, SHEETS  # noqa: E402
from kho.store import LocalStore, mo_store  # noqa: E402

DANH_MUC = ["KHO_MISA", "MA_MISA", "HANG_KHO", "QUY_DOI", "KHOI_LUONG_RIENG", "DONG_NGHIA_DVT", "DOI_TAC", "BO_PHAN",
            "TON_MISA_THAM_KHAO"]


def mo(args):
    if args.local:
        return LocalStore(args.local)
    if os.environ.get("GOOGLE_CREDENTIALS"):  # GitHub Actions
        import json
        return mo_store(dict(kieu="sheets", credentials=json.loads(os.environ["GOOGLE_CREDENTIALS"]),
                             sheet_id_danh_muc=os.environ.get("SHEET_ID_DANH_MUC") or os.environ.get("SHEET_ID"),
                             sheet_id_phat_sinh=os.environ.get("SHEET_ID_PHAT_SINH") or os.environ.get("SHEET_ID")))
    with open(os.path.join(ROOT, ".streamlit", "secrets.toml"), "rb") as f:
        cfg = tomllib.load(f)["luu_tru"]
    return mo_store(dict(cfg))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("lenh", choices=["khoi_tao", "mau_danh_muc", "nap_danh_muc", "sao_luu", "kiem_tra_ket_noi"])
    ap.add_argument("duong_dan", nargs="?")
    ap.add_argument("--local")
    a = ap.parse_args()

    if a.lenh == "mau_danh_muc":
        with pd.ExcelWriter(a.duong_dan) as w:
            for n in DANH_MUC:
                pd.DataFrame(columns=SHEETS[n]).to_excel(w, sheet_name=n, index=False)
        print("Đã tạo mẫu", a.duong_dan)
        return

    s = mo(a)
    if a.lenh == "kiem_tra_ket_noi":
        for d in s.kiem_tra():
            print(d)
        print("KẾT NỐI ĐẠT: app đọc và ghi được vào CSDL.")
    elif a.lenh == "khoi_tao":
        s.khoi_tao()
        print("Đã tạo/kiểm tra đủ sheet:", ", ".join(SHEETS))
    elif a.lenh == "nap_danh_muc":
        data = s.read_all()
        src = pd.read_excel(a.duong_dan, sheet_name=None, dtype=str)
        for n in DANH_MUC:
            if n not in src:
                continue
            df = src[n].fillna("")
            if list(df.columns[: len(SHEETS[n])]) != SHEETS[n]:
                raise SystemExit(f"Sheet {n}: tiêu đề sai. Cần {SHEETS[n]}")
            df = df[df.apply(lambda r: any(str(v).strip() for v in r), axis=1)]
            if df.empty:
                continue
            if not data[n].empty:
                print(f"BỎ QUA {n}: sheet đích đã có {len(data[n])} dòng (không nạp chồng).")
                continue
            for c in df.columns:  # cột số: đổi sang số (ô trống giữ trống), sai định dạng thì dừng, không đoán
                if c in NUMERIC:
                    def so(v, c=c):
                        v = str(v).strip().replace(",", ".")
                        try:
                            return "" if v == "" else float(v)
                        except ValueError:
                            raise SystemExit(f"Sheet {n}, cột {c}: giá trị “{v}” không phải số.")
                    df[c] = df[c].map(so)
            s.append(n, df.to_dict("records"))
            print(f"Đã nạp {n}: {len(df)} dòng")
    elif a.lenh == "sao_luu":
        os.makedirs(a.duong_dan, exist_ok=True)
        for n, df in s.read_all().items():
            df.to_csv(os.path.join(a.duong_dan, f"{n}.csv"), index=False, encoding="utf-8-sig")
        print("Đã sao lưu vào", a.duong_dan)


if __name__ == "__main__":
    main()
