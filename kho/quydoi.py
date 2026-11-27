"""Quy đổi kg ↔ tấm/cây theo KÍCH THƯỚC và khối lượng riêng (QT01).

Công thức và cách đọc kích thước trong tên hàng lấy nguyên từ _checkpoint/scripts_so_ton/hs.py
(đã dùng để dựng phiếu kiểm kê v3), để số liệu trong app khớp số đã duyệt.
Kích thước luôn tính bằng mm; khối lượng riêng (klr) bằng g/cm3; kết quả bằng kg trên 1 đơn vị (1 tấm / 1 cây).
"""
from __future__ import annotations

import math
import re

from .schema import DUNG_SAI

# Khối lượng riêng mặc định (QT01, người dùng đã duyệt). Sửa trong sheet KHOI_LUONG_RIENG (chỉ quản trị).
KLR_MAC_DINH = {
    "thép": 7.85, "inox 304": 7.93, "inox 201": 7.93, "inox 316": 7.98, "inox 430": 7.70,
    "nhôm": 2.70, "nhôm A5052": 2.68, "đồng": 8.93,
}

# Hình dạng -> (nhãn, [nhãn các kích thước, mm])
HINH_DANG = {
    "tấm": ("Tấm", ["Dày", "Rộng", "Dài"]),
    "ống": ("Ống tròn", ["Đường kính ngoài", "Dày thành", "Dài"]),
    "hộp": ("Hộp (vuông/chữ nhật)", ["Cạnh a", "Cạnh b", "Dày thành", "Dài"]),
    "láp": ("Láp / tròn đặc", ["Đường kính", "Dài"]),
    "thép V": ("Thép góc V", ["Cạnh a", "Cạnh b", "Dày", "Dài"]),
}


def kg_ly_thuyet(hinh_dang: str, klr: float, kt: list[float]) -> float:
    """kg lý thuyết của 1 tấm / 1 cây. kt: kích thước (mm) đúng thứ tự trong HINH_DANG."""
    if hinh_dang not in HINH_DANG:
        raise ValueError(f"Hình dạng không hỗ trợ: {hinh_dang}")
    n = len(HINH_DANG[hinh_dang][1])
    if len(kt) != n or any(not (float(x) > 0) for x in kt):
        raise ValueError("Phải nhập đủ các kích thước, mỗi số lớn hơn 0.")
    if not (float(klr) > 0):
        raise ValueError("Khối lượng riêng phải lớn hơn 0.")
    k = [float(x) for x in kt]
    if hinh_dang == "tấm":
        t, w, l = k
        return t * w * l * klr / 1e6
    if hinh_dang == "ống":
        d, t, l = k
        if t >= d / 2:
            raise ValueError("Dày thành phải nhỏ hơn nửa đường kính ngoài.")
        return math.pi * (d - t) * t * l * klr / 1e6
    if hinh_dang == "hộp":
        a, b, t, l = k
        if t >= min(a, b) / 2:
            raise ValueError("Dày thành phải nhỏ hơn nửa cạnh nhỏ.")
        return 2 * t * (a + b - 2 * t) * l * klr / 1e6   # chưa tính bo góc
    if hinh_dang == "láp":
        d, l = k
        return math.pi / 4 * d * d * l * klr / 1e6
    a, b, t, l = k                                        # thép V, chưa tính bo góc
    if t >= min(a, b) / 2:
        raise ValueError("Dày phải nhỏ hơn nửa cạnh nhỏ.")
    return t * (a + b - t) * l * klr / 1e6


def lech(a: float, b: float) -> float | None:
    """Độ lệch tương đối (a − b) / b. None nếu b không dương."""
    return None if not b or b <= 0 else (a - b) / b


def vuot_dung_sai(a: float, b: float) -> bool:
    d = lech(a, b)
    return d is not None and abs(d) > DUNG_SAI + 1e-12


def kg_tu_sl(sl: float, he_so: float) -> float:
    return sl * he_so


def sl_tu_kg(kg: float, he_so: float) -> float:
    return kg / he_so if he_so > 0 else 0.0


def tach_nguyen_le(kg: float, he_so: float) -> tuple[int, float]:
    """kg -> (số đơn vị nguyên, kg còn lẻ). Dùng khi hiện 'tương đương n tấm + x kg'."""
    n = int(kg // he_so) if he_so > 0 else 0
    return n, kg - n * he_so


# ---------------- Đọc vật liệu và kích thước từ tên hàng (nguyên từ hs.py) ----------------

def vat_lieu(name: str):
    """(nhãn vật liệu, klr) theo tên hàng; (None, None) nếu không phải kim loại tính được."""
    n = name.lower()
    if re.search(r'nhựa|mica|uhmw|\babs\b|\bpvc\b|\bpe\b|cao su|gỗ|giấy|lưới', n):
        return (None, None)
    if 'inox' in n:
        if re.search(r'\b430\b', n):
            return ('inox 430', KLR_MAC_DINH["inox 430"])
        m = re.search(r'\b(304|316|201)\b', n)
        if m and m.group(1) == '316':
            return ('inox 316', KLR_MAC_DINH["inox 316"])
        return ('inox ' + (m.group(1) if m else '304'), KLR_MAC_DINH["inox 304"])
    if re.search(r'nhôm', n) and not re.search(r'mạ nhôm|nhôm kẽm|alu ?zinc', n):
        if 'a5052' in n.replace(' ', ''):
            return ('nhôm A5052', KLR_MAC_DINH["nhôm A5052"])
        return ('nhôm', KLR_MAC_DINH["nhôm"])
    if 'đồng' in n:
        return ('đồng', KLR_MAC_DINH["đồng"])
    return ('thép', KLR_MAC_DINH["thép"])


def _prep(name: str) -> str:
    s = ' ' + name.lower() + ' '
    s = re.sub(r'\([^)]*(kg|c/|cái|cây|=|cắt|chặt|nhập)[^)]*\)', ' ', s)
    s = re.sub(r'(\d),(\d)', r'\1.\2', s)
    s = s.replace('×', 'x').replace('*', ' x ').replace('ф', 'ø').replace('φ', 'ø')
    s = re.sub(r'(\d)\s*x\s*(?=\d)', r'\1 x ', s)
    s = re.sub(r'(?<![a-zà-ỹ])(?:phi|pi|ø|d)\s*(?=\d)', ' ø ', s)
    s = re.sub(r'(\d)\s*(ly|li)\b', r'\1 ly ', s)
    s = re.sub(r'(\d)\s*mm\b', r'\1 mm ', s)
    s = re.sub(r'(\d)\s*m\b', r'\1 m ', s)
    s = re.sub(r'(?<![\d.])\b[a-z]+[a-z\-.]*\d[\w\-./]*', ' ', s)
    s = re.sub(r'\b(304|316|201|430)(/2b)?\b', ' ', s)
    s = re.sub(r'\s*x\s*', ' x ', s)
    return re.sub(r'\s+', ' ', s)


def _nums(s: str):
    return [(float(m.group(1)), m.group(2)) for m in re.finditer(r'(\d+(?:\.\d+)?)\s*(ly|mm|m)?\b', s)]


def _dai(v, u):
    if u == 'm':
        return v * 1000
    if u == 'mm':
        return v
    return v * 1000 if v < 20 else v


def doc_kich_thuoc(name: str):
    """Đọc hình dạng + kích thước (mm) từ tên hàng. Trả về dict(hinh_dang, kt=[...]) hoặc None nếu không chắc."""
    n = name.lower()
    if re.search(r'lưới|bending|chấn|chống trượt|chặt|sai khổ|nhựa|mica|uhmw|x c\s*$', n):
        return None
    if re.search(r'gân', n) and not re.search(r'phi gân', n):
        return None
    s = _prep(name)
    ns = _nums(s)
    if re.search(r'\btấm\b|thép mạ kẽm|thép cuộn', n):
        if len(ns) != 3:
            return None
        ly = [i for i, (v, u) in enumerate(ns) if u == 'ly']
        if len(ly) == 1:
            t = ns[ly[0]][0]
            rest = [v for i, (v, u) in enumerate(ns) if i != ly[0]]
        else:
            t = ns[0][0]
            rest = [ns[1][0], ns[2][0]]
        w, l = [x * 1000 if x < 10 else x for x in rest]
        if not (t <= 50 and t < w and t < l):
            return None
        return dict(hinh_dang='tấm', kt=[t, w, l])
    if re.search(r'hộp|vuông đen|nhôm .*vuông', n):
        if len(ns) != 4:
            return None
        ly = [i for i, (v, u) in enumerate(ns) if u == 'ly']
        if len(ly) == 1:
            t = ns[ly[0]][0]
            rest = [x for i, x in enumerate(ns) if i != ly[0]]
            a, b, (lv, lu) = rest[0][0], rest[1][0], rest[2]
        else:
            (a, _), (b, _), (t, _), (lv, lu) = ns
        l = _dai(lv, lu)
        if not (t < min(a, b) / 2 and l > 50):
            return None
        return dict(hinh_dang='hộp', kt=[a, b, t, l])
    if re.search(r'ống|tròn phi|nhôm .*tròn', n):
        if len(ns) != 3:
            return None
        ly = [i for i, (v, u) in enumerate(ns) if u == 'ly']
        if ly and ly != [1]:
            return None
        (d, _), (t, _), (lv, lu) = ns
        l = _dai(lv, lu)
        if not (t < d / 2 and l > 20):
            return None
        return dict(hinh_dang='ống', kt=[d, t, l])
    if re.search(r'láp|\blap\b|tròn đặc|thép tròn|phi gân|thép phi', n):
        if len(ns) != 2:
            return None
        (d, _), (lv, lu) = ns
        return dict(hinh_dang='láp', kt=[d, _dai(lv, lu)])
    if re.search(r'thép hình v|thép góc v|(^|\s)v\s*\.?\s*\d', n) and 'lệch' not in n:
        if len(ns) == 4:
            ly = [i for i, (v, u) in enumerate(ns) if u == 'ly']
            if len(ly) == 1:
                t = ns[ly[0]][0]
                rest = [x for i, x in enumerate(ns) if i != ly[0]]
                a, b, (lv, lu) = rest[0][0], rest[1][0], rest[2]
            else:
                (a, _), (b, _), (t, _), (lv, lu) = ns
        elif len(ns) == 3 and ns[1][1] == 'ly':
            a = ns[0][0]
            b = a
            t = ns[1][0]
            lv, lu = ns[2]
        else:
            return None
        if not (t < min(a, b) / 2):
            return None
        return dict(hinh_dang='thép V', kt=[a, b, t, _dai(lv, lu)])
    return None
