"""Điền thông tin tài khoản dịch vụ từ tệp JSON (tải từ Google Cloud) vào .streamlit/secrets.toml.

Không in khóa bí mật ra màn hình. Giữ nguyên 2 ID file Google Sheets đã có trong secrets.toml.

Dùng:  python scripts/nhap_khoa_json.py "C:\\Users\\...\\Downloads\\app-kho-516003-xxxx.json"
       python scripts/nhap_khoa_json.py            (tự tìm tệp app-kho-*.json mới nhất trong thư mục Downloads)
"""
import glob
import json
import os
import sys
import tomllib

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SECRETS = os.path.join(ROOT, ".streamlit", "secrets.toml")
CAN = ("type", "project_id", "private_key_id", "private_key", "client_email", "client_id")


def q(v) -> str:
    """Chuỗi TOML hợp lệ (json.dumps thoát đúng \\n, dấu nháy, ký tự lạ)."""
    return json.dumps(str(v), ensure_ascii=True)


def main():
    if len(sys.argv) > 1:
        duong_dan = sys.argv[1]
    else:
        tim = glob.glob(os.path.join(os.path.expanduser("~"), "Downloads", "app-kho-*.json"))
        if not tim:
            raise SystemExit("Không thấy tệp app-kho-*.json trong thư mục Downloads. Gõ đường dẫn tệp sau tên lệnh.")
        duong_dan = max(tim, key=os.path.getmtime)
    with open(duong_dan, encoding="utf-8") as f:
        k = json.load(f)
    thieu = [x for x in CAN if not k.get(x)]
    if thieu:
        raise SystemExit(f"Tệp JSON thiếu trường: {thieu}. Đây có đúng là khóa của tài khoản dịch vụ không?")
    if k["type"] != "service_account" or not k["private_key"].startswith("-----BEGIN PRIVATE KEY-----"):
        raise SystemExit("Tệp không đúng dạng khóa tài khoản dịch vụ (type phải là service_account, có private_key).")
    cu = {}
    if os.path.exists(SECRETS):
        with open(SECRETS, "rb") as f:
            cu = tomllib.load(f).get("luu_tru", {})
    id_dm, id_ps = cu.get("sheet_id_danh_muc", ""), cu.get("sheet_id_phat_sinh", "")
    id_n1, id_n2 = cu.get("sheet_id_nm1", ""), cu.get("sheet_id_nm2", "")
    if not id_dm or not ((id_n1 and id_n2) or id_ps):
        raise SystemExit("secrets.toml chưa có sheet_id_danh_muc và (sheet_id_nm1 + sheet_id_nm2, hoặc sheet_id_phat_sinh). "
                         "Điền ID các file trước rồi chạy lại.")
    noi_dung = (
        "# Tệp bí mật của app. KHÔNG đưa lên GitHub, không gửi cho ai.\n"
        "[luu_tru]\n"
        'kieu = "sheets"\n'
        f"sheet_id_danh_muc  = {q(id_dm)}   # APP_KHO_DANH_MUC\n"
        + (f"sheet_id_phat_sinh = {q(id_ps)}   # APP_KHO_PHAT_SINH_2026 (file cũ / chung)\n" if id_ps else "")
        + (f"sheet_id_nm1 = {q(id_n1)}   # APP_KHO_PHAT_SINH_NM1_2026\n" if id_n1 else "")
        + (f"sheet_id_nm2 = {q(id_n2)}   # APP_KHO_PHAT_SINH_NM2_2026\n" if id_n2 else "")
        + "\n"
        "[luu_tru.credentials]\n"
        + "".join(f"{x} = {q(k[x])}\n" for x in CAN)
        + f"token_uri = {q(k.get('token_uri', 'https://oauth2.googleapis.com/token'))}\n")
    with open(SECRETS, "w", encoding="utf-8", newline="\n") as f:
        f.write(noi_dung)
    with open(SECRETS, "rb") as f:   # đọc lại để chắc tệp hợp lệ
        kiem = tomllib.load(f)["luu_tru"]
    assert kiem["credentials"]["private_key"] == k["private_key"] and kiem["sheet_id_danh_muc"] == id_dm
    print("Đã điền khóa vào .streamlit/secrets.toml (không hiển thị nội dung bí mật).")
    print("  Tài khoản dịch vụ:", k["client_email"])
    print("  Dự án:", k["project_id"], "| mã khóa:", k["private_key_id"][:6] + "…")
    print("  File danh mục :", id_dm[:8] + "…", "| file phát sinh:", id_ps[:8] + "…")
    print("Tiếp theo: chuyển tệp JSON ra khỏi thư mục Downloads (xem hướng dẫn) để tránh lộ khóa.")


if __name__ == "__main__":
    main()
