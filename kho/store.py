"""Lớp lưu trữ: Google Sheets (chạy thật) hoặc file Excel (chạy thử trên máy).

Chỉ có 2 thao tác: đọc tất cả sheet, ghi thêm dòng vào 1 sheet.
"""
from __future__ import annotations

import os
import threading

import pandas as pd

from .schema import NHA_MAY, NUMERIC, SHEETS, SHEETS_DANH_MUC


def _to_frame(name: str, rows: list[list]) -> pd.DataFrame:
    cols = SHEETS[name]
    if not rows:
        return pd.DataFrame(columns=cols)
    header = [str(h).strip() for h in rows[0]]
    if header[: len(cols)] != cols:
        raise ValueError(f"Sheet {name}: dòng tiêu đề không đúng mẫu. Cần: {cols}; đang có: {header}")
    body = [list(r) + [""] * (len(cols) - len(r)) for r in rows[1:]]
    body = [r[: len(cols)] for r in body if any(str(c).strip() for c in r)]
    df = pd.DataFrame(body, columns=cols)
    for c in cols:
        if c in NUMERIC:
            df[c] = pd.to_numeric(df[c].replace("", None), errors="coerce")
        else:
            df[c] = df[c].map(lambda v: "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v).strip())
            # số nguyên đọc thành '12.0' -> '12'
            df[c] = df[c].str.replace(r"^(-?\d+)\.0$", r"\1", regex=True)
    return df


def _row_values(name: str, row: dict) -> list:
    out = []
    for c in SHEETS[name]:
        v = row.get(c, "")
        if v is None or (isinstance(v, float) and v != v):   # None hoặc NaN (ô số trống đọc từ bảng) -> ô trống
            v = ""
        if c in NUMERIC and v != "":
            v = float(v)
            if v.is_integer():
                v = int(v)
        elif c not in NUMERIC:
            v = str(v)
        out.append(v)
    return out


class LocalStore:
    """File Excel trên máy – chỉ để chạy thử, 1 người dùng."""

    _lock = threading.Lock()

    def __init__(self, path: str):
        self.path = path
        self.khoi_tao()   # tạo file / bổ sung sheet còn thiếu

    def khoi_tao(self):
        from openpyxl import Workbook, load_workbook
        with self._lock:
            if os.path.exists(self.path):
                wb = load_workbook(self.path)
            else:
                os.makedirs(os.path.dirname(os.path.abspath(self.path)), exist_ok=True)
                wb = Workbook()
                wb.remove(wb.active)
            for name, cols in SHEETS.items():
                if name not in wb.sheetnames:
                    wb.create_sheet(name).append(cols)
            wb.save(self.path)

    def mo_ta(self) -> list[dict]:
        return [dict(nhan="File Excel trên máy (chạy thử)", ten=os.path.basename(self.path), url=os.path.abspath(self.path))]

    def kiem_tra(self) -> list[str]:
        d = self.read_all()
        return [f"File {os.path.abspath(self.path)}: đọc được {len(d)} sheet – ĐẠT"]

    def read_all(self) -> dict[str, pd.DataFrame]:
        from openpyxl import load_workbook
        with self._lock:
            wb = load_workbook(self.path, read_only=True, data_only=True)
            data = {}
            for name in SHEETS:
                rows = [list(r) for r in wb[name].iter_rows(values_only=True)]
                rows = [["" if c is None else c for c in r] for r in rows]
                data[name] = _to_frame(name, rows)
            wb.close()
        return data

    def append(self, name: str, rows: list[dict]):
        from openpyxl import load_workbook
        if not rows:
            return
        with self._lock:
            wb = load_workbook(self.path)
            ws = wb[name]
            for r in rows:
                vals = _row_values(name, r)
                ws.append(vals)
                for cell in ws[ws.max_row]:  # giữ chuỗi như chuỗi, không thành công thức
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        cell.data_type = "s"
            wb.save(self.path)


class SheetsStore:
    """Google Sheets qua tài khoản dịch vụ (service account).

    spreadsheet_ids:
      - một chuỗi: 1 file chứa tất cả sheet;
      - {"danh_muc": id, "phat_sinh": id}: 2 file (phát sinh chung hai nhà máy – cách cũ);
      - {"danh_muc": id, "NM1": id, "NM2": id}: 3 file, dữ liệu phát sinh TÁCH THEO NHÀ MÁY
        (ghi vào file của nhà máy đó, đọc thì gộp lại).
    """

    def __init__(self, credentials_info: dict, spreadsheet_ids):
        import gspread
        if isinstance(spreadsheet_ids, str):
            spreadsheet_ids = {"danh_muc": spreadsheet_ids, "phat_sinh": spreadsheet_ids}
        self._gc = gspread.service_account_from_dict(credentials_info)
        self.tach_nha_may = all(nm in spreadsheet_ids for nm in NHA_MAY)
        keys = ["danh_muc"] + (list(NHA_MAY) if self.tach_nha_may else ["phat_sinh"])
        self._files = {}
        for k in keys:
            i = spreadsheet_ids[k]
            self._files[k] = next((f for kk, f in self._files.items() if spreadsheet_ids[kk] == i), None) \
                or self._gc.open_by_key(i)
        gop = not self.tach_nha_may and spreadsheet_ids["danh_muc"] == spreadsheet_ids["phat_sinh"]
        # nơi chứa từng sheet: khóa file, hoặc None = có mặt ở MỌI file nhà máy (chế độ tách)
        self._of = {n: ("danh_muc" if gop or n in SHEETS_DANH_MUC else ("phat_sinh" if not self.tach_nha_may else None))
                    for n in SHEETS}

    # ---- định tuyến ----
    def _khoa_cua_dong(self, name: str, row: dict) -> str:
        """File sẽ nhận 1 dòng của sheet `name`."""
        k = self._of[name]
        if k is not None:
            return k
        nm = str(row.get("nha_may", "") or "").strip()
        if not nm:   # HUY_PHIEU: suy ra từ số phiếu (NM2-N1-…)
            sp = str(row.get("so_phieu", "") or "")
            nm = sp[:3] if sp[:3] in NHA_MAY else ""
        if not nm and row.get("nguoi") == "HE_THONG":   # dòng thử kết nối
            nm = NHA_MAY[0]
        if nm not in NHA_MAY:
            raise ValueError(f"Không biết ghi dòng của sheet {name} vào file nhà máy nào (thiếu nha_may): {row}")
        return nm

    def _ten_theo_file(self):
        out = {}
        for n, k in self._of.items():
            for kk in ([k] if k is not None else list(NHA_MAY)):
                out.setdefault(kk, []).append(n)
        return out

    def khoi_tao(self):
        """Tạo các sheet còn thiếu kèm dòng tiêu đề. Không đụng sheet đã có dữ liệu."""
        for key, names in self._ten_theo_file().items():
            sh = self._files[key]
            have = {ws.title: ws for ws in sh.worksheets()}
            for name in names:
                cols = SHEETS[name]
                if name not in have:
                    ws = sh.add_worksheet(title=name, rows=1000, cols=len(cols))
                    ws.update([cols], "A1", value_input_option="RAW")
                    ws.freeze(rows=1)
                else:
                    first = have[name].row_values(1)
                    if not first:
                        have[name].update([cols], "A1", value_input_option="RAW")
                    elif first == cols:
                        pass
                    elif len(first) < len(cols) and first == cols[: len(first)]:
                        # bản app mới có thêm cột ở cuối: nối thêm tiêu đề, không đụng dữ liệu cũ
                        have[name].update([cols], "A1", value_input_option="RAW")
                    else:
                        raise ValueError(f"Sheet {name} đã có nhưng tiêu đề khác mẫu: {first}")

    def read_all(self, cho_phep_thieu: bool = False) -> dict[str, pd.DataFrame]:
        """Đọc mọi sheet. cho_phep_thieu=True: sheet chưa tồn tại coi như rỗng (chỉ dùng khi đọc file CŨ lúc chuyển dữ liệu)."""
        mot_file = {}   # sheet nằm ở 1 file -> bảng
        nhieu = {}      # sheet có ở mọi file nhà máy -> list bảng (theo thứ tự NM1, NM2)

        def nhan(name, df):
            if self._of[name] is None:
                nhieu.setdefault(name, []).append(df)
            else:
                mot_file[name] = df

        for key, names in self._ten_theo_file().items():
            if cho_phep_thieu:
                have = {ws.title for ws in self._files[key].worksheets()}
                for n in [n for n in names if n not in have]:
                    nhan(n, _to_frame(n, []))
                names = [n for n in names if n in have]
            if not names:
                continue
            res = self._files[key].values_batch_get(
                [f"'{n}'" for n in names],
                params={"valueRenderOption": "UNFORMATTED_VALUE", "dateTimeRenderOption": "FORMATTED_STRING"})
            for name, vr in zip(names, res.get("valueRanges", [])):
                nhan(name, _to_frame(name, vr.get("values", [])))
        for name, dfs in nhieu.items():
            mot_file[name] = pd.concat(dfs, ignore_index=True) if len(dfs) > 1 else dfs[0]
        return {n: mot_file[n] for n in SHEETS}

    def append(self, name: str, rows: list[dict]):
        if not rows:
            return
        theo_file = {}
        for r in rows:
            theo_file.setdefault(self._khoa_cua_dong(name, r), []).append(_row_values(name, r))
        # 1 lần gọi cho mỗi file = ghi trọn cả phiếu (1 phiếu luôn thuộc đúng 1 nhà máy)
        for key, values in theo_file.items():
            self._files[key].worksheet(name).append_rows(
                values, value_input_option="RAW", insert_data_option="INSERT_ROWS", table_range="A1")

    def mo_ta(self) -> list[dict]:
        """Nơi đang lưu dữ liệu (hiển thị trong app cho người quản trị yên tâm)."""
        out, thay = [], set()
        nhan_cua = {"danh_muc": "Danh mục chung (dữ liệu gốc, đối tác)", "phat_sinh": "Phát sinh & hệ thống",
                    **{nm: f"Phát sinh riêng {nm}" for nm in NHA_MAY}}
        gop = len(set(map(id, self._files.values()))) == 1
        for key, sh in self._files.items():
            if id(sh) in thay:
                continue
            thay.add(id(sh))
            out.append(dict(nhan="Google Sheets (cả danh mục lẫn phát sinh)" if gop else f"Google Sheets – {nhan_cua[key]}",
                            ten=sh.title, url=sh.url))
        return out

    def kiem_tra(self) -> list[str]:
        """Kiểm tra kết nối thật: liệt kê sheet, ghi 1 dòng thử vào NHAT_KY rồi đọc lại. Ném lỗi nếu sai."""
        import uuid
        dong = []
        for key, names in self._ten_theo_file().items():
            sh = self._files[key]
            have = {ws.title for ws in sh.worksheets()}
            thieu = [n for n in names if n not in have]
            dong.append(f"File “{sh.title}” ({sh.url}): có {len(names) - len(thieu)}/{len(names)} sheet cần dùng"
                        + (f" – THIẾU: {', '.join(thieu)} (chạy: python scripts/cong_cu.py khoi_tao)" if thieu else " – đủ"))
            if thieu:
                raise RuntimeError("\n".join(dong))
        dau = f"KIEM_TRA_{uuid.uuid4().hex[:8]}"
        cac_nm = list(NHA_MAY) if self.tach_nha_may else [None]
        for nm in cac_nm:   # ghi 1 dòng thử vào NHAT_KY của TỪNG file nhà máy
            self.append("NHAT_KY", [dict(thoi_diem="", nguoi="HE_THONG", hanh_dong="KIEM_TRA_KET_NOI",
                                         chi_tiet=f"{dau}{'_' + nm if nm else ''}", nha_may=nm or "")])
        doc = set(self.read_all()["NHAT_KY"]["chi_tiet"])
        thieu_doc = [nm for nm in cac_nm if f"{dau}{'_' + nm if nm else ''}" not in doc]
        if thieu_doc:
            raise RuntimeError("Đã ghi dòng thử vào NHAT_KY nhưng đọc lại không thấy"
                               f"{' (' + ', '.join(thieu_doc) + ')' if thieu_doc[0] else ''}. Kiểm tra quyền chia sẻ file.")
        dong.append(f"Ghi và đọc lại dòng thử trong NHAT_KY{' của cả ' + ' và '.join(NHA_MAY) if self.tach_nha_may else ''}: "
                    f"ĐẠT (mã {dau}). Bạn có thể xóa các dòng này trong Sheets.")
        return dong


def mo_store(cfg: dict):
    """Mở kho dữ liệu theo cấu hình.

    - Excel chạy thử: {'kieu': 'local', 'duong_dan': ...}
    - Google, 3 file tách nhà máy: {'kieu': 'sheets', 'sheet_id_danh_muc', 'sheet_id_nm1', 'sheet_id_nm2', 'credentials'}
    - Google, 2 file: {'kieu': 'sheets', 'sheet_id_danh_muc', 'sheet_id_phat_sinh', 'credentials'}
    - Google, 1 file cho tất cả: chỉ có 'sheet_id'.
    """
    if cfg.get("kieu") == "sheets":
        dm = cfg.get("sheet_id_danh_muc") or cfg.get("sheet_id")
        nm1, nm2 = cfg.get("sheet_id_nm1"), cfg.get("sheet_id_nm2")
        if bool(nm1) != bool(nm2):
            raise ValueError("Thiếu sheet_id_nm1 hoặc sheet_id_nm2 trong secrets (phải có đủ cả hai để tách theo nhà máy).")
        if nm1 and nm2:
            if not dm:
                raise ValueError("Thiếu sheet_id_danh_muc trong secrets.")
            if len({dm, nm1, nm2}) < 3:
                raise ValueError("sheet_id_danh_muc, sheet_id_nm1, sheet_id_nm2 phải là 3 file khác nhau.")
            return SheetsStore(dict(cfg["credentials"]), {"danh_muc": dm, "NM1": nm1, "NM2": nm2})
        ps = cfg.get("sheet_id_phat_sinh") or cfg.get("sheet_id")
        if not dm or not ps:
            raise ValueError("Thiếu sheet_id_danh_muc / sheet_id_phat_sinh (hoặc sheet_id) trong secrets.")
        return SheetsStore(dict(cfg["credentials"]), {"danh_muc": dm, "phat_sinh": ps})
    return LocalStore(cfg.get("duong_dan", "du_lieu_thu/APP_KHO_DATA_THU.xlsx"))
