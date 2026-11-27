"""APP KHO NM1 – NM2: nhập, xuất, tự tính tồn.

Chạy thử trên máy:  streamlit run app.py   (dùng file du_lieu_thu/APP_KHO_DATA_THU.xlsx)
Chạy thật: khai [luu_tru] trong secrets (xem README.md).

Giao diện thiết kế cho laptop 15 inch: mọi màn hình chính vừa 1 trang, không phải cuộn.
"""
from __future__ import annotations

import io
import os
from datetime import date, timedelta

import pandas as pd
import streamlit as st

from kho import logic as L
from kho import quydoi as Q
from kho import reports as R
from kho.auth import check_pin, hash_pin
from kho.schema import DUNG_SAI, LOAI, NHA_MAY, NHAN_O, QUYEN_QUAN_TRI
from kho.store import mo_store

import base64

TEN_APP = "APP KHO METALIC"
_THU_MUC_APP = os.path.dirname(os.path.abspath(__file__))
_LOGO = os.path.join(_THU_MUC_APP, "assets", "favicon.png")   # biểu tượng Metalic (lấy từ metalic.com.vn)

st.set_page_config(page_title=TEN_APP, page_icon=_LOGO if os.path.exists(_LOGO) else "📦",
                   layout="wide", initial_sidebar_state="collapsed")


def logo_html(cao: int) -> str:
    """Thẻ <img> logo (nhúng base64); không có file thì trả về rỗng."""
    try:
        with open(_LOGO, "rb") as f:
            b = base64.b64encode(f.read()).decode()
        return f'<img src="data:image/png;base64,{b}" style="height:{cao}px;vertical-align:middle;margin-right:8px">'
    except OSError:
        return ""

NGAY_MAC_DINH = 15      # bộ lọc phiếu: mặc định lấy 15 ngày gần nhất (hôm nay 27 → từ 12)
CAO_DS = 430            # chiều cao khung danh sách / chi tiết phiếu (px)
CAO_DONG = 262          # chiều cao khung "các dòng trong phiếu" (px)

st.markdown("""
<style>
/* ===== Khung chung: gọn, vừa 1 màn hình laptop ===== */
header[data-testid="stHeader"], [data-testid="stToolbar"], [data-testid="stDecoration"] { display: none !important; }
.stApp { background: #f3f5f8; }
.block-container { max-width: 1500px; padding: 0.7rem 1.6rem 0.6rem !important; }
[data-testid="stVerticalBlock"] { gap: 0.55rem; }
[data-testid="stHorizontalBlock"] { gap: 0.8rem; }
html, body, .stApp, .stMarkdown p, .stCaption p { font-size: 16.5px; }
h1, h2, h3, h4 { padding: 0 !important; margin: 0 !important; }
h4 { font-size: 18px !important; }
[data-testid="stWidgetLabel"] { min-height: 0; margin-bottom: 1px; }
[data-testid="stWidgetLabel"] p { font-size: 14px !important; font-weight: 600; color: #3b4452; }
input, textarea { font-size: 16.5px !important; }
div[data-baseweb="select"] > div, div[data-baseweb="base-input"], div[data-baseweb="input"] { min-height: 42px; }
li[role="option"], li[role="option"] div { font-size: 16.5px !important; }
[data-testid="stVerticalBlockBorderWrapper"] { background: #fff; border-radius: 10px; }
[data-testid="stVerticalBlockBorderWrapper"] > div { padding-top: 0.15rem; padding-bottom: 0.15rem; }
[data-testid="stAlert"] { padding: 0.45rem 0.8rem; }
[data-testid="stForm"] { padding: 0; }

/* ===== Nút ===== */
.stButton, .stDownloadButton, .stFormSubmitButton,
[data-testid="stElementContainer"]:has(> .stButton), [data-testid="stElementContainer"]:has(> .stDownloadButton),
[data-testid="stElementContainer"]:has(> .stFormSubmitButton) { width: 100% !important; }
.stButton button, .stDownloadButton button, .stFormSubmitButton button {
  min-height: 42px; width: 100% !important; border-radius: 8px; padding: 0.2rem 0.9rem; }
.stButton button p, .stDownloadButton button p, .stFormSubmitButton button p {
  font-size: 16.5px; font-weight: 600; white-space: normal !important; }
.stButton button[kind="primary"], .stFormSubmitButton button[kind="primary"] { box-shadow: 0 1px 2px rgba(0,0,0,.18); }

/* Thẻ chức năng trong tab */
[class*="st-key-the_"] button { min-height: 78px; justify-content: flex-start; text-align: left;
  padding: 8px 16px; border: 1.5px solid #cdd9e5; background: #fff; border-radius: 10px; }
[class*="st-key-the_"] button:hover { border-color: #1a7f37; background: #eef8f1; }
[class*="st-key-the_"] button p { font-size: 15px; font-weight: 500; line-height: 1.35; text-align: left; color: #4b5563; }
[class*="st-key-the_"] button p strong { color: #1a7f37; font-size: 19px; font-weight: 700; }
[class*="st-key-the_"] button div { justify-content: flex-start; }

/* Dòng trong danh sách phiếu (bấm để xem) */
[class*="st-key-pk_"] button { min-height: 38px; justify-content: flex-start; text-align: left;
  padding: 3px 12px; border-radius: 7px; border: 1px solid #e1e6ec; background: #fafbfc; }
[class*="st-key-pk_"] button:hover { border-color: #1a7f37; }
[class*="st-key-pk_"] button div, [class*="st-key-pk_"] button p { width: 100%; justify-content: flex-start; text-align: left; }
[class*="st-key-pk_"] button p { font-size: 15px; font-weight: 500; color: #1f2937; }
[class*="st-key-pk_"] button[kind="primary"] { background: #1a7f37; border-color: #1a7f37; }
[class*="st-key-pk_"] button[kind="primary"] p { font-weight: 700; color: #fff; }
[class*="st-key-xoa_"] button { min-height: 32px; padding: 0; color: #b42318; border-color: #f3c4c0; }

/* Thanh tab */
.st-key-tab_chinh button { min-height: 44px; }
.st-key-tab_chinh button p { font-size: 16.5px !important; font-weight: 700; }

/* Chữ tiêu đề */
.brand { font-size: 21px; font-weight: 800; color: #14532d; white-space: nowrap; }
.nguoi { font-size: 14.5px; color: #4b5563; text-align: right; line-height: 1.25; }
.tieu-de { font-size: 22px; font-weight: 700; color: #111827; }
.so-phieu { font-size: 26px; font-weight: 700; color: #1a7f37; padding: 6px 0; }
.nho { font-size: 14px; color: #6b7280; }
.ct-luoi { display: grid; grid-template-columns: repeat(3, 1fr); gap: 2px 18px; font-size: 15px; margin: 4px 0 8px; }
</style>""", unsafe_allow_html=True)


# ---------------- Dữ liệu ----------------

def _cau_hinh() -> dict:
    if os.environ.get("APP_KHO_FILE"):   # chủ động chạy thử trên file Excel: luôn ưu tiên, không đụng Google thật
        return {"kieu": "local", "duong_dan": os.environ["APP_KHO_FILE"]}
    try:
        if "luu_tru" in st.secrets:
            return dict(st.secrets["luu_tru"])
    except Exception:
        pass
    return {"kieu": "local", "duong_dan": os.environ.get("APP_KHO_FILE", "du_lieu_thu/APP_KHO_DATA_THU.xlsx")}


@st.cache_resource
def store():
    return mo_store(_cau_hinh())


@st.cache_data(ttl=60, show_spinner="Đang tải dữ liệu…")
def _doc():
    return store().read_all()


def du_lieu():
    return _doc()


def du_lieu_moi():
    """Đọc mới (bỏ bộ nhớ đệm) – dùng ngay trước khi ghi."""
    _doc.clear()
    return _doc()


def di(trang: str, **kw):
    st.session_state.trang = trang
    for k, v in kw.items():
        st.session_state[k] = v
    st.rerun()


def user() -> dict:
    return st.session_state.user


def la_qt() -> bool:
    return user()["quyen"] == QUYEN_QUAN_TRI


def _reset_phieu():
    for k in ("dong", "phieu_id", "ket_qua", "ket_ngay", "dem_dong", "chon_phieu"):
        st.session_state.pop(k, None)


def tieu_de_trang(tieu_de: str, ve: tuple = ("chu", {}), cot=(1.25, 5.75, 3.0)):
    """Hàng đầu của màn hình con: [⬅ Quay lại]  Tiêu đề  ......  [chỗ đặt nút bên phải]. Trả về cột bên phải."""
    c = st.columns(list(cot), vertical_alignment="center")
    if c[0].button("⬅ Quay lại", key=f"ve_{ve[0]}_{ve[1].get('loai', '')}"):
        _reset_phieu()
        di(ve[0], **ve[1])
    c[1].markdown(f'<div class="tieu-de">{tieu_de}</div>', unsafe_allow_html=True)
    return c[2]


# ---------------- Đăng nhập ----------------

def trang_dang_nhap():
    _, giua, _ = st.columns([1, 1.2, 1])
    with giua:
        st.markdown(f'<div class="brand" style="font-size:28px;margin-top:8vh">{logo_html(48)}{TEN_APP}</div>'
                    '<div class="nho" style="margin-bottom:8px">Kho NM1 – NM2 · nhập, xuất, tồn kho</div>',
                    unsafe_allow_html=True)
        _dang_nhap()


def _dang_nhap():
    data = du_lieu()
    nd = L.nguoi_dung(data)
    nd = nd[nd["hoat_dong"].map(lambda v: str(v).strip() not in ("", "0"))]
    if nd.empty:
        st.warning("Chưa có người dùng. Tạo tài khoản QUẢN TRỊ đầu tiên.")
        with st.form("tao_qt"):
            ma = st.text_input("Mã người dùng (viết tắt, không dấu, vd: TU)").strip().upper()
            ten = st.text_input("Họ tên")
            nm = st.selectbox("Nhà máy", NHA_MAY)
            pin = st.text_input("Mã PIN (4–8 số)", type="password")
            if st.form_submit_button("TẠO QUẢN TRỊ", type="primary"):
                if not ma or not ten or not (pin.isdigit() and 4 <= len(pin) <= 8):
                    st.error("Điền đủ mã, họ tên và PIN 4–8 số.")
                elif not L.nguoi_dung(du_lieu_moi()).empty:
                    st.error("Đã có người dùng. Tải lại trang.")
                else:
                    store().append("NGUOI_DUNG", [dict(ma=ma, ten=ten, nha_may=nm, quyen=QUYEN_QUAN_TRI,
                                                       pin_hash=hash_pin(pin), hoat_dong="1",
                                                       cap_nhat_luc=L.bay_gio().strftime("%Y-%m-%d %H:%M:%S"))])
                    du_lieu_moi()
                    st.success("Đã tạo. Đăng nhập bên dưới.")
                    st.rerun()
        return
    sai = st.session_state.get("so_lan_sai", 0)
    if sai >= 5:
        st.error("Nhập sai PIN quá 5 lần. Tải lại trang (F5) và thử lại.")
        return
    with st.form("dang_nhap"):
        ten = st.selectbox("Bạn là ai?", nd["ten"].tolist(), index=None, placeholder="Chọn tên của bạn")
        pin = st.text_input("Mã PIN", type="password")
        bam = st.form_submit_button("VÀO", type="primary")
    if bam:
        r = nd[nd["ten"] == ten]
        if r.empty or not check_pin(pin, r.iloc[0]["pin_hash"]):
            st.session_state.so_lan_sai = sai + 1
            st.error("Sai tên hoặc PIN.")
        else:
            st.session_state.so_lan_sai = 0
            st.session_state.user = r.iloc[0].to_dict()
            L.ghi_nhat_ky(store(), r.iloc[0]["ma"], "DANG_NHAP", "")
            di("chu")


# ---------------- Thanh đầu trang + tab ----------------

TAB_NHAP, TAB_XUAT, TAB_DB, TAB_XEM, TAB_QT = "📥 NHẬP", "📤 XUẤT", "🔁 ĐẶC BIỆT", "📋 TỒN & PHIẾU", "🛠 QUẢN TRỊ"

MO_TA = {
    "N1": "Hàng mua từ nhà cung cấp về kho", "N2": "Xưởng trả lại vật tư lĩnh dư",
    "N3": "Hàng gửi gia công ngoài đã về", "N4": "Nhận hàng nhà máy kia chuyển sang",
    "N5": "Phần dư sau cắt, phế liệu thu hồi", "X1": "Giao vật tư cho xưởng – ghi LSX, PO",
    "X2": "Vật tư dùng chung cho tổ / bộ phận", "X3": "Gửi hàng ra ngoài gia công",
    "X4": "Trả hàng lỗi / sai cho NCC", "X5": "Chuyển hàng sang nhà máy kia",
    "X6": "Hàng hỏng, mất phẩm chất", "X7": "Xuất phế liệu đem bán",
    "X8": "1 tấm lớn → tấm nhỏ + phần dư", "X9": "Chuyển hàng giữa sổ MTL và CMA",
}


def ds_tab():
    return [TAB_NHAP, TAB_XUAT, TAB_DB, TAB_XEM] + ([TAB_QT] if la_qt() else [])


def _doi_tab():
    _reset_phieu()
    st.session_state.trang = "chu"
    if st.session_state.get("tab_chinh") is None:  # bấm lại tab đang mở: giữ tab cũ
        st.session_state.tab_chinh = st.session_state.get("tab_truoc", TAB_NHAP)


def dau_trang():
    """Một hàng duy nhất: tên app | các tab | người dùng + đăng xuất."""
    u = user()
    if "tab_muon" in st.session_state:  # chuyển tab theo lệnh
        st.session_state.tab_chinh = st.session_state.pop("tab_muon")
    if st.session_state.get("tab_chinh") not in ds_tab():
        st.session_state.tab_chinh = TAB_NHAP
    st.session_state.tab_truoc = st.session_state.tab_chinh
    c = st.columns([2.1, 6.4, 2.5], vertical_alignment="center")
    c[0].markdown(f'<div class="brand">{logo_html(30)}{TEN_APP}</div>', unsafe_allow_html=True)
    with c[1]:
        st.segmented_control("Chức năng", ds_tab(), key="tab_chinh", on_change=_doi_tab,
                             label_visibility="collapsed", width="stretch")
    with c[2]:
        cc = st.columns([2.3, 1.2], vertical_alignment="center")
        cc[0].markdown(f'<div class="nguoi">👤 <b>{u["ten"]}</b><br>{u["nha_may"]}'
                       f'{" · Quản trị" if la_qt() else ""} · {L.hom_nay():%d/%m/%Y}</div>', unsafe_allow_html=True)
        if cc[1].button("Thoát", key="dang_xuat"):
            st.session_state.clear()
            st.rerun()


def _luoi_the(muc: list[tuple], so_cot=4):
    """muc: (trang đích, tiêu đề, mô tả, tham số)"""
    cols = st.columns(so_cot, gap="small")
    for i, (dich, td, mt, kw) in enumerate(muc):
        if cols[i % so_cot].button(f"**{td}**  \n{mt}", key=f"the_{dich}_{kw.get('loai', '')}"):
            _reset_phieu()
            di(dich, **kw)


def trang_chu():
    tab = st.session_state.tab_chinh
    u = user()
    data = du_lieu()
    if tab == TAB_NHAP:
        cho = L.x5_cho_nhan(data, u["nha_may"])["phieu_id"].nunique()
        if cho:
            st.info(f"🚚 Có {cho} phiếu chuyển kho đang chờ {u['nha_may']} nhận – bấm thẻ N4.")
        muc = [("ds_loai", f"{m} · {LOAI[m]['ten']}", MO_TA[m], dict(loai=m)) for m in ["N1", "N2", "N3"]]
        muc.append(("ds_loai", f"N4 · {LOAI['N4']['ten']}" + (f" ({cho} chờ)" if cho else ""), MO_TA["N4"], dict(loai="N4")))
        muc.append(("ds_loai", f"N5 · {LOAI['N5']['ten']}", MO_TA["N5"], dict(loai="N5")))
        _luoi_the(muc)
    elif tab == TAB_XUAT:
        _luoi_the([("ds_loai", f"{m} · {LOAI[m]['ten']}", MO_TA[m], dict(loai=m))
                   for m in ["X1", "X2", "X3", "X4", "X5", "X6", "X7"]])
    elif tab == TAB_DB:
        _luoi_the([("ds_loai", f"{m} · {LOAI[m]['ten']}", MO_TA[m], dict(loai=m)) for m in ["X8", "X9"]])
    elif tab == TAB_XEM:
        _luoi_the([("ton", "XEM TỒN", "Tồn hiện tại theo nhà máy, tìm theo tên", {}),
                   ("ds_phieu", "PHIẾU ĐÃ LẬP", "Xem lại, xem chi tiết, hủy phiếu (mọi loại)", {}),
                   ("quy_doi", "MÁY TÍNH QUY ĐỔI", "Đổi qua lại giữa số tấm / cây và kg", {})])
    elif tab == TAB_QT:
        _luoi_the([("bao_cao", "BÁO CÁO EXCEL", "Nhập · Xuất · Nhập-xuất-tồn cho kế toán", {}),
                   ("kiem_ke", "KIỂM KÊ", "Đưa số kiểm thực tế lên, tạo phiếu K1", {}),
                   ("khoa_so", "KHÓA SỔ", "Chốt kỳ, chụp tồn cuối kỳ", {}),
                   ("dm_quan_tri", "DANH MỤC & QUY ĐỔI", "Thêm hàng, hệ số theo kích thước, khối lượng riêng", {}),
                   ("danh_muc", "KIỂM TRA DANH MỤC", "Mã thiếu, hệ số chưa duyệt, kg hóa đơn lệch…", {}),
                   ("nguoi_dung", "NGƯỜI DÙNG", "Thêm người, đổi PIN, khóa tài khoản", {})])
        noi_luu = []
        for x in store().mo_ta():
            ten = f'<a href="{x["url"]}" target="_blank">{x["ten"]}</a>' if str(x["url"]).startswith("http") else x["ten"]
            noi_luu.append(f'<b>{x["nhan"]}</b>: {ten}')
        st.markdown('<div class="nho">💾 Mọi thao tác ghi thẳng vào nơi lưu dữ liệu: ' + " · ".join(noi_luu) + "</div>",
                    unsafe_allow_html=True)


# ---------------- Lập phiếu (vừa 1 màn hình) ----------------

def _chon_hang(data, nha_may, so, xuat: bool, key: str):
    hang = L.hang_hoat_dong(data)
    hang = hang[hang["so"] == so]
    ton = L.tinh_ton(data, nha_may).set_index("ma_hang")["ton_kho"]
    hang = hang.assign(ton=hang["ma_hang"].map(ton).fillna(0.0))
    if xuat:
        hang = hang[hang["ton"] > 1e-9]
    tk = st.text_input("🔍 Tìm hàng (gõ vài chữ, không cần dấu, vd: thep tam 3)", key=f"tk_{key}")
    if len(tk.strip()) < 2:
        st.markdown('<div class="nho">👆 Gõ vài chữ vào ô trên để tìm hàng, rồi mới chọn ở đây.</div>',
                    unsafe_allow_html=True)
        st.selectbox("Chọn hàng", [], index=None, placeholder="(gõ ô tìm hàng trước)", disabled=True,
                     key=f"sel_{key}_trong")
        return None
    kq = L.tim_hang(hang, tk).sort_values("ten_hang")
    if kq.empty:
        st.warning("Không thấy hàng." + (" (Chỉ hiện hàng còn tồn.)" if xuat else ""))
        return None
    if len(kq) > 200:
        st.caption(f"Có {len(kq)} hàng – gõ thêm chữ để thu hẹp. Đang hiện 200 hàng đầu.")
        kq = kq.head(200)
    nhan = {r["ma_hang"]: f"{r['ten_hang']}  |  {r['dvt_kho']}  |  tồn: {r['ton']:g}" for _, r in kq.iterrows()}
    ma = st.selectbox("Chọn hàng", list(nhan), format_func=nhan.get, index=None,
                      placeholder="Bấm để chọn", key=f"sel_{key}")
    return None if ma is None else kq[kq["ma_hang"] == ma].iloc[0]


class _Luoi:
    """Xếp các ô nhập thành lưới nhiều cột, ô nào đến lượt thì lấy cột kế tiếp."""

    def __init__(self, so_cot=4):
        self.k, self.n, self.cols = so_cot, 0, None

    def o(self):
        if self.n % self.k == 0:
            self.cols = st.columns(self.k)
        c = self.cols[self.n % self.k]
        self.n += 1
        return c


def _man_hinh_da_luu(loai):
    sp = st.session_state.ket_qua
    st.markdown(f'<div class="so-phieu">✅ ĐÃ LƯU PHIẾU: {sp}</div>', unsafe_allow_html=True)
    c = st.columns(3)
    if c[0].button("➕ LẬP THÊM PHIẾU", type="primary"):
        _reset_phieu()
        st.rerun()
    if c[1].button("👁 XEM PHIẾU VỪA LẬP"):
        ngay = st.session_state.get("ket_ngay")
        _reset_phieu()
        _loc_xem_phieu(loai, sp, ngay)
        di("ds_loai", loai=loai)
    if c[2].button(f"📋 VỀ DANH SÁCH {loai}"):
        _reset_phieu()
        di("ds_loai", loai=loai)


def trang_phieu():
    loai = st.session_state.loai
    cfg = LOAI[loai]
    u = user()
    data = du_lieu()
    cot_phai = tieu_de_trang(f"{loai} · {cfg['ten']}", ve=("ds_loai", {"loai": loai}), cot=(1.25, 6.25, 2.5))
    if st.session_state.get("ket_qua"):
        _man_hinh_da_luu(loai)
        return
    st.session_state.setdefault("dong", [])
    st.session_state.setdefault("phieu_id", L.PhieuNhap(loai, "", "").phieu_id)
    ds = st.session_state.dong
    khoa_dau = bool(ds)

    # ---- Thông tin phiếu (lưới 4 cột) ----
    g = _Luoi(4)
    with st.container(border=True):
        nha_may = g.o().selectbox("Nhà máy", NHA_MAY, index=NHA_MAY.index(u["nha_may"]),
                                  disabled=not la_qt() or khoa_dau)
        ngay = g.o().date_input("Ngày phiếu", value=L.hom_nay(), max_value=L.hom_nay(), format="DD/MM/YYYY")
        cac_so = L.so_cua_nha_may(data, nha_may)
        if not cac_so:
            st.error(f"{nha_may} chưa khai kho MISA (sheet KHO_MISA). Báo quản trị.")
            return
        if loai == "X9":
            if len(cac_so) < 2:
                st.error(f"{nha_may} chỉ có 1 sổ – không chuyển sổ được.")
                return
            chieu = g.o().selectbox("Chiều chuyển sổ", [f"{cac_so[0]} → {cac_so[1]}", f"{cac_so[1]} → {cac_so[0]}"],
                                    disabled=khoa_dau)
            so_x, so_n = chieu.split(" → ")
        else:
            so = g.o().selectbox("Công ty (sổ)", cac_so, disabled=khoa_dau or len(cac_so) == 1)
            so_x = so_n = so
        gt = {}
        for o in [o for o in cfg["bat_buoc"] + cfg["hien"] if o != "phieu_lien_quan"]:
            nhan = NHAN_O[o] + (" *" if o in cfg["bat_buoc"] else "")
            if o == "doi_tac":
                dt = data["DOI_TAC"]
                dt = dt[dt["loai"] == cfg.get("doi_tac")]["ten_doi_tac"].drop_duplicates().tolist()
                gt[o] = g.o().selectbox(nhan, dt, index=None, placeholder="Chọn hoặc gõ tên mới rồi Enter",
                                        accept_new_options=True, help="Chọn trong danh sách, hoặc gõ tên mới rồi nhấn Enter. "
                                        "Tên mới sẽ được lưu vào danh sách để lần sau chọn lại.") or ""
            elif o == "bo_phan":
                bp = data["BO_PHAN"]
                bp = bp[bp["nha_may"] == nha_may]["ten_bo_phan"].drop_duplicates().tolist()
                gt[o] = g.o().selectbox(nhan, bp, index=None, placeholder="Chọn hoặc gõ tên rồi Enter",
                                        accept_new_options=True) or ""
            elif o == "kho_den":
                gt[o] = g.o().selectbox(nhan, [n for n in NHA_MAY if n != nha_may])
            else:
                gt[o] = g.o().text_input(nhan)
        ghi_chu_phieu = g.o().text_input("Ghi chú phiếu")

    # ---- Trái: thêm hàng · Phải: các dòng đã thêm ----
    trai, phai = st.columns([5.6, 6.4], gap="medium")
    n = st.session_state.get("dem_dong", 0)
    with trai:
        with st.container(border=True):
            dau = cfg["dau"]
            if dau == 0:
                lua = st.radio("Dòng này là", ["XUẤT ra (hàng đem đi / đem cắt)", "NHẬP vào (hàng nhận về / sau cắt)"],
                               horizontal=True, key=f"chieu_{n}")
                dau = -1 if lua.startswith("XUẤT") else 1
            h = _chon_hang(data, nha_may, so_x if dau < 0 else so_n, xuat=dau < 0, key=str(n))
            hs = L.he_so(data, h["ma_hang"], str(ngay)) if h is not None else None
            # Ô "kg hóa đơn" chỉ có nghĩa với phiếu nhập mua và hàng mà kế toán ghi theo kg (thép tấm, ống, hộp…)
            la_kg = loai == "N1" and hs is not None and str(hs["dvt_misa"]).strip().lower() == "kg"
            c = st.columns([1.1, 1.3, 0.9, 1.2] if la_kg else [1.3, 1.1, 1.6])
            sl = c[0].number_input(f"Số lượng ({h['dvt_kho']})" if h is not None else "Số lượng",
                                   min_value=0.0, step=1.0, format="%.2f", key=f"sl_{n}")
            kg_hd = 0.0
            if la_kg:
                kg_hd = c[1].number_input(
                    "Kg ghi trên hóa đơn (nếu có)", min_value=0.0, step=1.0, format="%.3f", key=f"kg_{n}",
                    help="Số kg mà nhà cung cấp ghi trên HÓA ĐƠN (hoặc phiếu cân) cho dòng hàng này. Để 0 nếu chưa có. "
                         "App so với kg quy đổi từ số tấm/cây và cảnh báo nếu lệch quá ±2%.")
            lo = c[-2].text_input("Lô (nếu có)", key=f"lo_{n}")
            gc = c[-1].text_input("Ghi chú dòng", key=f"gc_{n}")
            c = st.columns([2.2, 3], vertical_alignment="center")
            if h is not None and hs is None:
                c[1].error("Chưa có hệ số quy đổi đã duyệt – báo quản trị.")
            elif h is not None:
                he = float(hs["he_so"])
                txt = f'1 {h["dvt_kho"]} = <b>{he:g} {hs["dvt_misa"]}</b>'
                if sl > 0:
                    txt += f' · {sl:g} {h["dvt_kho"]} ≈ <b>{sl * he:,.3f} {hs["dvt_misa"]}</b>'
                pct = L.lech_kg_hoa_don(sl * he, hs["dvt_misa"], kg_hd) if sl > 0 else None
                if pct is not None:
                    vuot = abs(pct) > DUNG_SAI + 1e-12
                    txt += (f'<br><span style="color:{"#b42318" if vuot else "#1a7f37"}"><b>Hóa đơn lệch {pct * 100:+.2f}%'
                            f'{f" ⚠ VƯỢT ±{DUNG_SAI * 100:g}%" if vuot else " (trong dung sai)"}</b></span>')
                c[1].markdown(f'<div class="nho">{txt}</div>', unsafe_allow_html=True)
            if c[0].button("➕ THÊM DÒNG VÀO PHIẾU", disabled=h is None or hs is None):
                if sl <= 0:
                    st.error("Số lượng phải lớn hơn 0.")
                else:
                    ds.append(dict(ma_hang=h["ma_hang"], ten=h["ten_hang"], dvt=h["dvt_kho"], sl=sl, dau=dau,
                                   lo=lo, gc=gc, qd=sl * float(hs["he_so"]), dvt_misa=hs["dvt_misa"], kg_hd=kg_hd,
                                   pct=L.lech_kg_hoa_don(sl * float(hs["he_so"]), hs["dvt_misa"], kg_hd)))
                    st.session_state.dem_dong = n + 1
                    st.rerun()
    with phai:
        with st.container(border=True, height=CAO_DONG):
            if not ds:
                st.markdown('<div class="nho">Chưa có dòng nào. Chọn hàng ở khung bên trái, nhập số lượng rồi bấm '
                            '<b>THÊM DÒNG VÀO PHIẾU</b>.</div>', unsafe_allow_html=True)
            else:
                st.markdown(f"**Các dòng trong phiếu ({len(ds)})**")
                for i, d in enumerate(ds):
                    r = st.columns([0.45, 3.6, 1.3, 1.8, 0.6], vertical_alignment="center")
                    chieu_txt = ("▲ nhập " if d["dau"] > 0 else "▼ xuất ") if cfg["dau"] == 0 else ""
                    r[0].markdown(f"{i + 1}")
                    r[1].markdown(f"{chieu_txt}{d['ten']}" + (f"  <span class='nho'>lô {d['lo']}</span>" if d["lo"] else ""),
                                  unsafe_allow_html=True)
                    r[2].markdown(f"**{d['sl']:g}** {d['dvt']}")
                    hd = ""
                    if d.get("kg_hd"):
                        pct = d.get("pct")
                        vuot = pct is not None and abs(pct) > DUNG_SAI + 1e-12
                        mau = "#b42318" if vuot else "#1a7f37"
                        phan_tram = "" if pct is None else f" ({pct * 100:+.1f}%)"
                        hd = f"<br><span style='color:{mau}'>HĐ {d['kg_hd']:g} kg{phan_tram}{' ⚠' if vuot else ''}</span>"
                    r[3].markdown(f"<span class='nho'>{d['qd']:,.1f} {d['dvt_misa']}{hd}</span>", unsafe_allow_html=True)
                    if r[4].button("✕", key=f"xoa_{i}", help="Xóa dòng này"):
                        ds.pop(i)
                        st.rerun()

    # ---- Nút LƯU: nằm trên đầu trang, luôn nhìn thấy ----
    with cot_phai:
        luu = st.button("💾 LƯU PHIẾU", type="primary", disabled=not ds, key="luu_phieu")
    if luu:
        p = L.PhieuNhap(loai=loai, ngay=str(ngay), nha_may=nha_may, phieu_id=st.session_state.phieu_id,
                        ghi_chu_phieu=ghi_chu_phieu,
                        dong=[L.DongNhap(d["ma_hang"], d["sl"], dau=d["dau"], lo=d["lo"], ghi_chu=d["gc"],
                                         kg_hoa_don=d.get("kg_hd", 0.0)) for d in ds],
                        **gt)
        try:
            so_phieu = L.lap_phieu(store(), du_lieu_moi(), u, p)
        except L.LoiNghiepVu as e:
            st.error(f"⚠ {e}")
        except Exception as e:  # mạng/Google lỗi: giữ nguyên phiếu để bấm lại
            st.error(f"Lỗi kết nối, phiếu CHƯA chắc đã lưu. Bấm LƯU PHIẾU lại (không bị ghi trùng). Chi tiết: {e}")
        else:
            du_lieu_moi()
            st.session_state.ket_qua = so_phieu
            st.session_state.ket_ngay = str(ngay)
            st.rerun()


# ---------------- N4 nhận chuyển kho ----------------

def trang_nhan_ck():
    u = user()
    cot = tieu_de_trang("N4 · Nhận hàng chuyển kho", ve=("ds_loai", {"loai": "N4"}), cot=(1.25, 5.75, 3.0))
    nm = cot.selectbox("Nhà máy nhận", NHA_MAY, index=NHA_MAY.index(u["nha_may"]), disabled=not la_qt(),
                       label_visibility="collapsed")
    data = du_lieu()
    x5 = L.x5_cho_nhan(data, nm)
    if x5.empty:
        st.success("Không có phiếu chờ nhận.")
        return
    with st.container(border=False, height=CAO_DS + 60):
        for pid, g in x5.groupby("phieu_id", sort=False):
            d0 = g.iloc[0]
            with st.container(border=True):
                a, b = st.columns([8, 2], vertical_alignment="center")
                a.markdown(f"**{d0['so_phieu']}** · từ {d0['nha_may']} · ngày {d0['ngay']} · người lập {d0['nguoi_lap']}")
                if b.button("✅ ĐÃ NHẬN ĐỦ", key=f"nhan_{pid}", type="primary"):
                    try:
                        dm = du_lieu_moi()
                        p = L.phieu_nhan_tu_x5(dm, pid, str(L.hom_nay()))
                        p.phieu_id = f"N4-{pid}"  # cố định: bấm 2 lần không ghi trùng
                        sp = L.lap_phieu(store(), dm, u, p)
                    except L.LoiNghiepVu as e:
                        st.error(f"⚠ {e}")
                    else:
                        du_lieu_moi()
                        st.success(f"Đã lưu phiếu nhận {sp}")
                        st.rerun()
                st.dataframe(g[["ten_hang", "sl_kho", "dvt_kho", "lo"]].rename(
                    columns={"ten_hang": "Tên hàng", "sl_kho": "SL", "dvt_kho": "ĐVT", "lo": "Lô"}),
                    hide_index=True, width="stretch", height=min(35 + 35 * len(g), 220))
                st.caption("Nhận thiếu/thừa: vẫn bấm nhận đủ, ghi chú và báo quản trị điều chỉnh.")


# ---------------- Xem tồn ----------------

def trang_ton():
    u = user()
    tieu_de_trang("Tồn kho hiện tại")
    data = du_lieu()
    c = st.columns([1.2, 1.6, 4])
    nm = c[0].selectbox("Nhà máy", NHA_MAY, index=NHA_MAY.index(u["nha_may"]))
    chi_con = c[1].selectbox("Hiển thị", ["Hàng còn tồn", "Tất cả hàng"]) == "Hàng còn tồn"
    tim = c[2].text_input("🔍 Tìm hàng (không cần dấu)")
    hang = L.hang_hoat_dong(data)
    hang = hang[hang["so"].isin(L.so_cua_nha_may(data, nm))]
    t = L.tinh_ton(data, nm).set_index("ma_hang")
    hang = hang.assign(ton=hang["ma_hang"].map(t["ton_kho"]).fillna(0.0),
                       ton_misa=hang["ma_hang"].map(t["ton_misa"]).fillna(0.0))
    hang["dvt_misa"] = hang["ma_hang"].map(L.he_so_bang(data, str(L.hom_nay()))["dvt_misa"]).fillna("")
    if chi_con:
        hang = hang[hang["ton"].abs() > 1e-9]
    hang = L.tim_hang(hang, tim)
    st.markdown(f'<div class="nho">{len(hang)} mặt hàng · tồn của {nm} tính đến hiện tại</div>', unsafe_allow_html=True)
    st.dataframe(hang[["ten_hang", "ton", "dvt_kho", "so", "ma_misa", "ton_misa", "dvt_misa", "ma_hang"]].rename(columns={
        "ten_hang": "Tên hàng", "ton": "Tồn", "dvt_kho": "ĐVT", "so": "Sổ", "ma_misa": "Mã kế toán",
        "ton_misa": "Tồn quy đổi", "dvt_misa": "ĐVT MISA", "ma_hang": "Mã quy cách"}).sort_values("Tên hàng"),
        hide_index=True, width="stretch", height=CAO_DS + 40)


# ---------------- Danh sách phiếu + chi tiết + hủy ----------------

def _loc_xem_phieu(loai: str | None, so_phieu: str, ngay: str | None):
    """Chuẩn bị bộ lọc để phiếu vừa lập chắc chắn hiện trong danh sách (kể cả khi lập lùi ngày)."""
    kp = loai or "tat_ca"
    hn = L.hom_nay()
    d_tu = hn - timedelta(days=NGAY_MAC_DINH)
    if ngay:
        try:
            d_tu = min(d_tu, date.fromisoformat(ngay))
        except ValueError:
            pass
    st.session_state[f"mdtu_{kp}"] = d_tu
    st.session_state[f"mdnm_{kp}"] = so_phieu.split("-")[0] if so_phieu.split("-")[0] in NHA_MAY else user()["nha_may"]
    st.session_state[f"ver_{kp}"] = st.session_state.get(f"ver_{kp}", 0) + 1
    st.session_state.chon_phieu = so_phieu


def _chi_tiet_phieu(data, r, u):
    g = data["GIAO_DICH"]
    g = g[g["phieu_id"] == r["phieu_id"]].sort_values("dong")
    d0 = g.iloc[0]
    tt = "🔴 ĐÃ HỦY" if r["da_huy"] else "🟢 Hiệu lực"
    st.markdown(f"#### {r['so_phieu']}  <span class='nho'>· {d0['loai']} {d0['ten_loai']} · {tt}</span>",
                unsafe_allow_html=True)
    ttin = [("Ngày", d0["ngay"]), ("Nhà máy", d0["nha_may"]), ("Sổ", ", ".join(g["so"].unique())),
            ("Người lập", f"{d0['nguoi_lap']} lúc {d0['thoi_diem_lap'][11:16]}")]
    for o in ("doi_tac", "bo_phan", "lsx", "po_kh", "so_chung_tu", "kho_den", "ly_do", "ghi_chu_phieu"):
        if str(d0[o]).strip():
            ttin.append((NHAN_O.get(o, "Ghi chú phiếu"), d0[o]))
    st.markdown('<div class="ct-luoi">' + "".join(f"<div>{k}: <b>{v}</b></div>" for k, v in ttin) + "</div>",
                unsafe_allow_html=True)
    cot_ct = ["dong", "chieu", "ten_hang", "sl_kho", "dvt_kho", "sl_misa", "dvt_misa", "kg_hoa_don", "lo", "ghi_chu_dong"]
    if not (g["kg_hoa_don"] > 0).any():
        cot_ct.remove("kg_hoa_don")
    st.dataframe(g.assign(chieu=g["dau"].map({1: "NHẬP", -1: "XUẤT"}))[cot_ct].rename(
        columns={"dong": "STT", "chieu": "Chiều", "ten_hang": "Tên hàng", "sl_kho": "SL", "dvt_kho": "ĐVT",
                 "sl_misa": "SL quy đổi", "dvt_misa": "ĐVT MISA", "kg_hoa_don": "Kg hóa đơn", "lo": "Lô",
                 "ghi_chu_dong": "Ghi chú"}),
        hide_index=True, width="stretch", height=min(40 + 35 * len(g), 230))
    if r["da_huy"]:
        h = data["HUY_PHIEU"]
        h = h[h["phieu_id"] == r["phieu_id"]].iloc[0]
        st.warning(f"Đã hủy bởi {h['nguoi_huy']} lúc {h['thoi_diem']}. Lý do: {h['ly_do']}")
        return
    with st.expander("Hủy phiếu này"):
        c = st.columns([3, 1], vertical_alignment="bottom")
        ly_do = c[0].text_input("Lý do hủy *", key=f"ly_do_{r['phieu_id']}")
        if c[1].button("HỦY PHIẾU", type="primary", key=f"huy_{r['phieu_id']}"):
            try:
                L.huy_phieu(store(), du_lieu_moi(), u, r["phieu_id"], ly_do)
            except L.LoiNghiepVu as e:
                st.error(f"⚠ {e}")
            else:
                du_lieu_moi()
                st.success("Đã hủy.")
                st.rerun()


def _bo_loc(kp: str):
    """Hàng lọc: Nhà máy | Từ ngày | Đến ngày | Người lập | Tìm. Đổi giá trị / nhấn Enter là lọc lại ngay."""
    u = user()
    hn = L.hom_nay()
    ver = st.session_state.get(f"ver_{kp}", 0)
    nm_mac_dinh = st.session_state.get(f"mdnm_{kp}", u["nha_may"])
    tu_mac_dinh = st.session_state.get(f"mdtu_{kp}", hn - timedelta(days=NGAY_MAC_DINH))
    # Trong biểu mẫu (form): nhấn Enter ở bất kỳ ô nào, hoặc bấm LỌC, là lọc lại ngay theo các giá trị vừa nhập.
    with st.form(f"loc_{kp}_{ver}", border=False):
        c = st.columns([1.0, 1.25, 1.25, 1.2, 2.3, 0.95], vertical_alignment="bottom")
        nm = c[0].selectbox("Nhà máy", NHA_MAY, index=NHA_MAY.index(nm_mac_dinh))
        tu_txt = c[1].text_input("Từ ngày (ngày/tháng/năm)", value=f"{tu_mac_dinh:%d/%m/%Y}", placeholder="vd 12/11/2026")
        den_txt = c[2].text_input("Đến ngày (ngày/tháng/năm)", value=f"{hn:%d/%m/%Y}", placeholder="vd 27/11/2026")
        nguoi = c[3].selectbox("Người lập", ["Của tôi", "Tất cả"], index=0 if not la_qt() else 1)
        tim = c[4].text_input("Tìm số phiếu / LSX / PO / đối tác")
        c[5].form_submit_button("🔍 LỌC", type="primary")
    tu, den = L.doc_ngay(tu_txt, hn.year), L.doc_ngay(den_txt, hn.year)
    return nm, tu, den, nguoi == "Của tôi", tim


def _danh_sach_phieu(loai: str | None):
    """Bộ lọc + danh sách bấm-để-xem (trái) + chi tiết phiếu (phải). loai=None: mọi loại."""
    u = user()
    data = du_lieu()
    kp = loai or "tat_ca"
    nm, tu, den, chi_toi, tim = _bo_loc(kp)
    if tu is None or den is None:
        st.warning("Chưa đọc được ngày. Gõ theo dạng ngày/tháng/năm, ví dụ 12/11/2026, rồi nhấn Enter.")
        return
    if tu > den:
        st.error("“Từ ngày” phải trước hoặc bằng “Đến ngày”.")
        return
    tat_ca = L.phieu_tong_hop(data)
    ds = tat_ca[(tat_ca["nha_may"] == nm) & (tat_ca["ngay"] >= str(tu)) & (tat_ca["ngay"] <= str(den))]
    if loai:
        ds = ds[ds["loai"] == loai]
    if chi_toi:
        ds = ds[ds["nguoi_lap"] == u["ma"]]
    if tim.strip():
        chuoi = (ds["so_phieu"] + " " + ds["lsx"] + " " + ds["po_kh"] + " " + ds["doi_tac"]).map(L.bo_dau)
        ds = ds[chuoi.str.contains(L.bo_dau(tim.strip()), regex=False)]
    ds = ds.sort_values(["ngay", "thoi_diem_lap"], ascending=False).reset_index(drop=True)  # ngày mới nhất trước

    trai, phai = st.columns([4.3, 7.7], gap="medium")
    chon = st.session_state.get("chon_phieu")
    if chon not in set(ds["so_phieu"]):
        chon = ds.iloc[0]["so_phieu"] if len(ds) else None  # mặc định: phiếu mới nhất
    with trai:
        st.markdown(f'<div class="nho"><b>{len(ds)} phiếu</b> · từ {tu:%d/%m/%Y} đến {den:%d/%m/%Y}'
                    f'{" · bấm một dòng để xem" if len(ds) else ""}</div>', unsafe_allow_html=True)
        with st.container(border=True, height=CAO_DS):
            if ds.empty:
                st.markdown('<div class="nho">Không có phiếu nào trong khoảng ngày đã chọn.'
                            + (" Bấm <b>THÊM MỚI</b> để lập phiếu." if loai else "") + "</div>", unsafe_allow_html=True)
            for i, r in ds.head(300).iterrows():
                nhan = f"{r['so_phieu']}  ·  {r['ngay'][8:10]}/{r['ngay'][5:7]}" + ("  ·  ĐÃ HỦY" if r["da_huy"] else "")
                if st.button(nhan, key=f"pk_{kp}_{i}", type="primary" if r["so_phieu"] == chon else "secondary"):
                    st.session_state.chon_phieu = r["so_phieu"]
                    st.rerun()
            if len(ds) > 300:
                st.caption(f"Chỉ hiện 300 phiếu mới nhất trong {len(ds)} phiếu – thu hẹp khoảng ngày để xem tiếp.")
    with phai:
        st.markdown('<div class="nho">Chi tiết phiếu</div>', unsafe_allow_html=True)
        with st.container(border=True, height=CAO_DS):
            if chon is None:
                st.markdown('<div class="nho">Chưa có phiếu để xem.</div>', unsafe_allow_html=True)
            else:
                _chi_tiet_phieu(data, tat_ca[tat_ca["so_phieu"] == chon].iloc[0], u)


def trang_ds_phieu():
    """Tab TỒN & PHIẾU: mọi loại phiếu."""
    tieu_de_trang("Phiếu đã lập (mọi loại)")
    _danh_sach_phieu(None)


def trang_ds_loai():
    """Bấm thẻ nghiệp vụ → danh sách phiếu của loại đó + nút THÊM MỚI."""
    loai = st.session_state.loai
    u = user()
    cot = tieu_de_trang(f"{loai} · {LOAI[loai]['ten']}")
    if loai == "N4":
        cho = L.x5_cho_nhan(du_lieu(), u["nha_may"])["phieu_id"].nunique()
        if cot.button(f"➕ NHẬN HÀNG ({cho} chờ)", type="primary", disabled=cho == 0, key="them_moi"):
            _reset_phieu()
            di("nhan_ck")
    else:
        if cot.button("➕ THÊM MỚI", type="primary", key="them_moi"):
            _reset_phieu()
            di("phieu", loai=loai)
    _danh_sach_phieu(loai)


# ---------------- Quản trị ----------------

def _chi_qt() -> bool:
    if not la_qt():
        st.error("Chỉ quản trị.")
        return False
    return True


def trang_bao_cao():
    tieu_de_trang("Báo cáo Excel cho kế toán")
    if not _chi_qt():
        return
    c = st.columns([1.3, 1.3, 1.3, 1.7], vertical_alignment="bottom")
    hn = L.hom_nay()
    tu = c[0].date_input("Từ ngày", value=hn.replace(day=1), format="DD/MM/YYYY")
    den = c[1].date_input("Đến ngày", value=hn, format="DD/MM/YYYY")
    nm = c[2].selectbox("Nhà máy", ["Cả NM1 + NM2"] + NHA_MAY)
    nm = None if nm.startswith("Cả") else nm
    if tu > den:
        st.error("Từ ngày phải trước đến ngày.")
        return
    tham_so = (str(tu), str(den), nm)
    if c[3].button("📊 TẠO BÁO CÁO", type="primary"):   # chỉ dựng file khi bấm; đọc dữ liệu mới nhất lúc bấm
        with st.spinner("Đang tạo 3 báo cáo…"):
            data = du_lieu_moi()
            st.session_state.bc = dict(tham_so=tham_so, files=dict(
                nhap=R.bao_cao_nhap(data, *tham_so), xuat=R.bao_cao_xuat(data, *tham_so), nxt=R.bao_cao_nxt(data, *tham_so)),
                luc=L.bay_gio().strftime("%H:%M:%S"))
    bc = st.session_state.get("bc")
    if bc and bc["tham_so"] == tham_so:
        duoi = f"{nm or 'NM1_NM2'}_{tu:%Y%m%d}_{den:%Y%m%d}.xlsx"
        xl = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        f = bc["files"]
        d = st.columns(3)
        d[0].download_button(f"⬇ BÁO CÁO NHẬP ({len(f['nhap']) // 1024} KB)", f["nhap"], f"BC_NHAP_{duoi}", xl)
        d[1].download_button(f"⬇ BÁO CÁO XUẤT ({len(f['xuat']) // 1024} KB)", f["xuat"], f"BC_XUAT_{duoi}", xl)
        d[2].download_button(f"⬇ NHẬP–XUẤT–TỒN ({len(f['nxt']) // 1024} KB)", f["nxt"], f"BC_NXT_{duoi}", xl)
        st.caption(f"Dữ liệu tại {bc['luc']}. Phiếu đã hủy không tính. Mỗi file tách sheet theo sổ MTL / CMA; "
                   "có sheet tổng hợp theo mã kế toán.")
    elif bc:
        st.info("Đã đổi điều kiện lọc – bấm TẠO BÁO CÁO để tạo lại.")
    else:
        st.markdown('<div class="nho">Chọn khoảng ngày và nhà máy rồi bấm <b>TẠO BÁO CÁO</b>.</div>', unsafe_allow_html=True)


def trang_kiem_ke():
    tieu_de_trang("Kiểm kê – đưa số kiểm thực tế lên app")
    if not _chi_qt():
        return
    st.markdown('<div class="nho">1) Tải mẫu → 2) điền cột <b>sl_thuc_te</b> (ô trống = không kiểm, giữ nguyên) → '
                '3) tải lên → 4) xem chênh lệch → 5) ghi phiếu K1.</div>', unsafe_allow_html=True)
    c = st.columns([1, 1.6, 1.6, 3], vertical_alignment="bottom")
    nm = c[0].selectbox("Nhà máy", NHA_MAY)
    ngay = str(c[1].date_input("Ngày kiểm (tồn đến hết ngày)", value=L.hom_nay(),
                               max_value=L.hom_nay(), format="DD/MM/YYYY"))
    data = du_lieu()
    mau = L.bang_kiem_ke(data, nm, ngay)
    bio = io.BytesIO()
    mau.to_excel(bio, index=False, sheet_name="KIEM_KE")
    c[2].download_button("⬇ TẢI MẪU KIỂM KÊ", bio.getvalue(), f"MAU_KIEM_KE_{nm}_{ngay}.xlsx")
    f = c[3].file_uploader("Tải lên file kiểm kê đã điền", type=["xlsx"], label_visibility="collapsed")
    if not f:
        return
    try:
        dem = pd.read_excel(f, dtype={"ma_hang": str})
        phieus, bang = L.phieu_kiem_ke(data, nm, ngay, dem)
    except L.LoiNghiepVu as e:
        st.error(f"⚠ {e}")
        return
    except Exception as e:
        st.error(f"File không đúng mẫu: {e}")
        return
    st.markdown(f'Số mã đã kiểm: **{len(bang)}** · số mã lệch: **{int((bang["chenh_lech"].abs() > 1e-9).sum())}**')
    st.dataframe(bang[["ma_hang", "ten_hang", "so", "ton_so_app", "sl_thuc_te", "chenh_lech"]].rename(columns={
        "ma_hang": "Mã", "ten_hang": "Tên", "so": "Sổ", "ton_so_app": "Tồn sổ app", "sl_thuc_te": "Thực tế",
        "chenh_lech": "Chênh lệch"}), hide_index=True, width="stretch", height=300)
    if not phieus:
        st.success("Không có chênh lệch – không cần ghi phiếu.")
        return
    c = st.columns([4, 2], vertical_alignment="center")
    xac_nhan = c[0].checkbox("Tôi đã kiểm tra bảng chênh lệch và duyệt điều chỉnh")
    if c[1].button("GHI PHIẾU KIỂM KÊ", type="primary", disabled=not xac_nhan):
        key = f"kk_{nm}_{ngay}_{f.name}_{f.size}"
        try:
            ra = []
            for i, p in enumerate(phieus):
                p.phieu_id = f"{key}_{i}"
                ra.append(L.lap_phieu(store(), du_lieu_moi(), user(), p))
        except L.LoiNghiepVu as e:
            st.error(f"⚠ {e}")
        else:
            du_lieu_moi()
            st.success("Đã ghi: " + ", ".join(ra))


def trang_khoa_so():
    tieu_de_trang("Khóa sổ")
    if not _chi_qt():
        return
    data = du_lieu()
    st.markdown("  ·  ".join(f"**{nm}**: khóa đến {L.khoa_den(data, nm) or '(chưa khóa)'}" for nm in NHA_MAY))
    c = st.columns([1, 1.6, 3])
    nm = c[0].selectbox("Nhà máy", NHA_MAY)
    den = c[1].date_input("Khóa đến hết ngày", value=L.hom_nay() - timedelta(days=1),
                          max_value=L.hom_nay(), format="DD/MM/YYYY")
    st.warning("Sau khi khóa: không ai lập/hủy phiếu có ngày ≤ ngày khóa. Không mở khóa trong app được.")
    c = st.columns([3, 2], vertical_alignment="center")
    ok = c[0].checkbox("Tôi chắc chắn")
    if c[1].button("KHÓA SỔ", type="primary", disabled=not ok):
        try:
            L.khoa_so(store(), du_lieu_moi(), user(), nm, str(den))
        except L.LoiNghiepVu as e:
            st.error(f"⚠ {e}")
        else:
            du_lieu_moi()
            st.success(f"Đã khóa {nm} đến {den} và chụp tồn vào sheet TON_CHOT.")


def trang_danh_muc():
    cot = tieu_de_trang("Kiểm tra danh mục")
    if not _chi_qt():
        return
    if cot.button("🔄 TẢI LẠI DỮ LIỆU"):
        du_lieu_moi()
    data = du_lieu()
    loi = L.kiem_tra_danh_muc(data)
    st.markdown(f'<div class="nho">Danh mục (HANG_KHO, MA_MISA, QUY_DOI, DOI_TAC, BO_PHAN, KHO_MISA) sửa trực tiếp trong '
                f'Google Sheets, sửa xong bấm Tải lại. · Hàng đang dùng: <b>{len(L.hang_hoat_dong(data))}</b> · '
                f'số lỗi: <b>{len(loi)}</b></div>', unsafe_allow_html=True)
    if loi.empty:
        st.success("Danh mục không có lỗi.")
    else:
        st.dataframe(loi, hide_index=True, width="stretch", height=260)
    cb = L.canh_bao_kg_hoa_don(data)
    st.markdown(f"**⚠ Phiếu nhập mua có kg hóa đơn lệch quá ±{DUNG_SAI * 100:g}% so với kg quy đổi: {len(cb)} dòng**")
    if cb.empty:
        st.success("Không có dòng nào vượt dung sai.")
    else:
        st.dataframe(cb.rename(columns={"so_phieu": "Số phiếu", "ngay": "Ngày", "ten_hang": "Tên hàng", "sl_kho": "SL",
                                        "dvt_kho": "ĐVT", "sl_misa": "Kg quy đổi", "kg_hoa_don": "Kg hóa đơn",
                                        "lech_pct": "Lệch %"}), hide_index=True, width="stretch", height=220)


def trang_nguoi_dung():
    tieu_de_trang("Người dùng")
    if not _chi_qt():
        return
    data = du_lieu()
    nd = L.nguoi_dung(data)
    st.dataframe(nd[["ma", "ten", "nha_may", "quyen", "hoat_dong", "cap_nhat_luc"]].rename(columns={
        "ma": "Mã", "ten": "Họ tên", "nha_may": "Nhà máy", "quyen": "Quyền", "hoat_dong": "Đang dùng",
        "cap_nhat_luc": "Cập nhật"}), hide_index=True, width="stretch", height=230)
    st.markdown("**Thêm / sửa người dùng** (đổi PIN, khóa tài khoản)")
    with st.form("nd"):
        c = st.columns([1, 2, 1, 1.2, 2, 1.3, 1], vertical_alignment="bottom")
        ma = c[0].text_input("Mã").strip().upper()
        ten = c[1].text_input("Họ tên")
        nm = c[2].selectbox("Nhà máy", NHA_MAY)
        quyen = c[3].selectbox("Quyền", ["KHO", QUYEN_QUAN_TRI])
        pin = c[4].text_input("PIN mới (trống = giữ cũ)", type="password")
        hd = c[5].selectbox("Trạng thái", ["Đang dùng", "Khóa"])
        luu = c[6].form_submit_button("LƯU", type="primary")
        if luu:
            cu = nd[nd["ma"] == ma]
            if not ma or not ma.isalnum() or not ten:
                st.error("Mã (chữ/số, không dấu) và họ tên bắt buộc.")
            elif pin and not (pin.isdigit() and 4 <= len(pin) <= 8):
                st.error("PIN phải 4–8 số.")
            elif cu.empty and not pin:
                st.error("Người mới phải có PIN.")
            elif ma == user()["ma"] and (hd == "Khóa" or quyen != QUYEN_QUAN_TRI):
                st.error("Không tự khóa / tự hạ quyền của chính mình.")
            else:
                store().append("NGUOI_DUNG", [dict(ma=ma, ten=ten, nha_may=nm, quyen=quyen,
                                                   pin_hash=hash_pin(pin) if pin else cu.iloc[0]["pin_hash"],
                                                   hoat_dong="1" if hd == "Đang dùng" else "0",
                                                   cap_nhat_luc=L.bay_gio().strftime("%Y-%m-%d %H:%M:%S"))])
                L.ghi_nhat_ky(store(), user()["ma"], "NGUOI_DUNG", f"{ma} {quyen} {hd}")
                du_lieu_moi()
                st.success("Đã lưu.")
                st.rerun()


# ---------------- Máy tính quy đổi (mọi người dùng) ----------------

def _chon_hang_chung(data, key: str, nhan_tim="🔍 Tìm hàng (không cần dấu)"):
    """Chọn 1 hàng bất kỳ trong danh mục. Trả về dòng hàng hoặc None."""
    hang = L.hang_hoat_dong(data)
    tk = st.text_input(nhan_tim, key=f"tk_{key}")
    if len(tk.strip()) < 2:
        st.markdown('<div class="nho">👆 Gõ vài chữ vào ô trên để tìm hàng, rồi mới chọn ở đây.</div>',
                    unsafe_allow_html=True)
        st.selectbox("Chọn hàng", [], index=None, placeholder="(gõ ô tìm hàng trước)", disabled=True,
                     key=f"sel_{key}_trong")
        return None
    kq = L.tim_hang(hang, tk).sort_values("ten_hang")
    if len(kq) > 200:
        st.caption(f"Có {len(kq)} hàng – gõ thêm chữ để thu hẹp (đang hiện 200 hàng đầu).")
        kq = kq.head(200)
    nhan = {r["ma_hang"]: f"{r['ten_hang']}  |  {r['dvt_kho']}  |  {r['so']} {r['ma_misa']}" for _, r in kq.iterrows()}
    ma = st.selectbox("Chọn hàng", list(nhan), format_func=nhan.get, index=None, placeholder="Bấm để chọn", key=f"sel_{key}")
    return None if ma is None else kq[kq["ma_hang"] == ma].iloc[0]


def _thong_tin_he_so(hs):
    kt = [hs[k] for k in ("kt_a", "kt_b", "kt_c", "kt_d") if pd.notna(hs[k])]
    d = [f"**Hệ số hiện hành:** 1 {hs['dvt_kho']} = **{float(hs['he_so']):g} {hs['dvt_misa']}**",
         f"Cách tính: {hs['cach_tinh'] or '—'}"]
    if hs["hinh_dang"] and kt:
        klr_txt = "" if pd.isna(hs["klr"]) else f" {float(hs['klr']):g} g/cm³"
        d.append(f"Kích thước: {hs['hinh_dang']} {' × '.join(f'{x:g}' for x in kt)} mm · {hs['vat_lieu']}{klr_txt}")
    if pd.notna(hs["kg_ly_thuyet"]) and hs["kg_ly_thuyet"] > 0:
        d.append(f"kg lý thuyết: {float(hs['kg_ly_thuyet']):.3f}")
    d.append(f"Duyệt: {hs['nguoi_duyet'] or '—'} ({hs['ngay_duyet'] or '—'})")
    return "  \n".join(d)


def trang_quy_doi():
    tieu_de_trang("Máy tính quy đổi: số tấm / cây ↔ kg")
    data = du_lieu()
    trai, phai = st.columns([5.5, 6.5], gap="medium")
    with trai:
        with st.container(border=True):
            h = _chon_hang_chung(data, "qd")
            hs = L.he_so(data, h["ma_hang"], str(L.hom_nay())) if h is not None else None
            if h is not None and hs is None:
                st.error("Hàng này chưa có hệ số quy đổi đã duyệt.")
            elif hs is not None:
                st.markdown(_thong_tin_he_so(hs))
    with phai:
        with st.container(border=True):
            if hs is None:
                st.markdown('<div class="nho">Chọn một hàng ở khung bên trái để quy đổi. Hệ số do quản trị duyệt, ở đây chỉ để xem '
                            'và tính.</div>', unsafe_allow_html=True)
            else:
                he, dk, dm = float(hs["he_so"]), hs["dvt_kho"], hs["dvt_misa"]
                st.markdown(f"**① Biết số {dk} → ra {dm}**")
                a = st.number_input(f"Số lượng ({dk})", min_value=0.0, step=1.0, format="%.2f", key="qd_sl")
                st.markdown(f"### {a:g} {dk} = {a * he:,.3f} {dm}")
                st.divider()
                st.markdown(f"**② Biết số {dm} → ra {dk}**")
                b = st.number_input(f"Số {dm}", min_value=0.0, step=1.0, format="%.3f", key="qd_kg")
                nguyen, le = Q.tach_nguyen_le(b, he)
                st.markdown(f"### {b:g} {dm} = {Q.sl_tu_kg(b, he):,.4f} {dk}")
                if b > 0:
                    st.markdown(f'<div class="nho">Tức {nguyen} {dk} nguyên và còn lẻ {le:,.3f} {dm}.</div>', unsafe_allow_html=True)


# ---------------- Danh mục & quy đổi (CHỈ QUẢN TRỊ) ----------------

CACH_NHAN = {"KICH_THUOC": "Tính theo kích thước (kg)", "CUNG_DVT": "Cùng đơn vị (hệ số 1)",
             "TRUC_TIEP": "Nhập hệ số trực tiếp (theo cân / chứng từ)"}


def _o_quy_doi(data, key: str, dvt_kho: str, dvt_misa: str, mac_dinh: dict | None = None):
    """Các ô nhập cách tính hệ số. Trả về dict tham số cho L.them_hang / L.doi_he_so (không gồm tên, ĐVT)."""
    dvt_kho, dvt_misa = dvt_kho.strip().lower(), dvt_misa.strip().lower()
    cach_ok = []
    if dvt_misa == "kg":
        cach_ok.append("KICH_THUOC")
    if dvt_kho and dvt_kho == dvt_misa:
        cach_ok.append("CUNG_DVT")
    cach_ok.append("TRUC_TIEP")
    cach = st.radio("Cách tính hệ số", cach_ok, format_func=CACH_NHAN.get, horizontal=True, key=f"cach_{key}")
    kw: dict = dict(cach=cach)
    if cach == "KICH_THUOC":
        klr = L.klr_hien_hanh(data)
        c = st.columns([1.2, 1.2, 0.9])
        hd = c[0].selectbox("Hình dạng", list(Q.HINH_DANG), format_func=lambda k: Q.HINH_DANG[k][0], key=f"hd_{key}")
        vl = c[1].selectbox("Vật liệu", list(klr), key=f"vl_{key}")
        c[2].markdown(f'<div class="nho" style="padding-top:1.7rem">KLR <b>{klr[vl]:g}</b> g/cm³</div>', unsafe_allow_html=True)
        nhan = Q.HINH_DANG[hd][1]
        cols = st.columns(len(nhan) + 1)
        kt = [cols[i].number_input(f"{t} (mm)", min_value=0.0, step=1.0, format="%.2f", key=f"kt{i}_{key}_{hd}")
              for i, t in enumerate(nhan)]
        kg_can = cols[-1].number_input(f"kg cân thực tế / 1 {dvt_kho or 'đơn vị'} (nếu có)", min_value=0.0, step=0.1,
                                       format="%.3f", key=f"kgc_{key}")
        kw.update(hinh_dang=hd, vat_lieu=vl, kt=kt, kg_can=kg_can)
        if all(x > 0 for x in kt):
            try:
                lt = Q.kg_ly_thuyet(hd, klr[vl], kt)
                st.markdown(f"**kg lý thuyết trên 1 {dvt_kho or 'đơn vị'}: {lt:,.4f} kg**" +
                            (f"  ·  hệ số sẽ dùng: **{kg_can:g} kg** (cân thực tế)" if kg_can > 0 else "  ·  hệ số sẽ dùng: kg lý thuyết"))
                if kg_can > 0 and Q.vuot_dung_sai(kg_can, lt):
                    st.warning(f"⚠ kg cân lệch {Q.lech(kg_can, lt) * 100:+.1f}% so với kg lý thuyết "
                               f"(vượt ±{DUNG_SAI * 100:g}%). Kiểm tra lại số cân hoặc kích thước.")
            except ValueError as e:
                st.error(str(e))
    elif cach == "TRUC_TIEP":
        c = st.columns([1.2, 2.8])
        kw["he_so_nhap"] = c[0].number_input(f"1 {dvt_kho or 'ĐVT kho'} = ? {dvt_misa or 'ĐVT MISA'}", min_value=0.0, step=0.1,
                                             format="%.6f", key=f"hsn_{key}")
        c[1].markdown('<div class="nho" style="padding-top:1.7rem">Bắt buộc ghi nguồn (số cân / chứng từ) vào ô ghi chú bên dưới.</div>',
                      unsafe_allow_html=True)
    return kw


def trang_dm_quan_tri():
    tieu_de_trang("Danh mục hàng & quy đổi (quản trị)")
    if not _chi_qt():
        return
    data = du_lieu()
    t1, t2, t3 = st.tabs(["➕ Thêm hàng mới", "✏️ Đổi hệ số hàng đã có", "⚖️ Khối lượng riêng"])

    with t1:
        c = st.columns([0.8, 3.2, 1.6])
        so = c[0].selectbox("Sổ", ["MTL", "CMA"], key="th_so")
        m = L.ma_misa_cua_so(data, so)
        nhan_m = {r["ma_misa"]: f"{r['ma_misa']} · {r['ten_misa']} · {r['dvt_misa']}" for _, r in m.iterrows()}
        ma_misa = c[1].selectbox("Mã kế toán (MISA) – gõ để tìm", list(nhan_m), format_func=nhan_m.get, index=None,
                                 placeholder="Gõ mã hoặc tên", key="th_ma")
        dvt_misa = m[m["ma_misa"] == ma_misa].iloc[0]["dvt_misa"] if ma_misa else ""
        c[2].markdown(f'<div class="nho" style="padding-top:1.7rem">ĐVT MISA: <b>{dvt_misa or "—"}</b></div>', unsafe_allow_html=True)
        c = st.columns([3.2, 1.2, 1.6, 1.6])
        ten = c[0].text_input("Tên hàng (quy cách) như kho gọi", key="th_ten")
        dvt_kho = c[1].text_input("ĐVT kho (tấm, cây, cái…)", key="th_dvt")
        nhom_ds = sorted(set(data["HANG_KHO"]["nhom"]) - {""})
        nv = c[2].selectbox("Nhóm hàng", nhom_ds + ["➕ Nhóm mới"], index=None, placeholder="Chọn", key="th_nhom")
        nhom = c[3].text_input("Tên nhóm mới", key="th_nhom_moi") if nv == "➕ Nhóm mới" else (nv or "")
        if ma_misa and dvt_kho.strip():
            kw = _o_quy_doi(data, "th", dvt_kho, dvt_misa)
            ghi_chu = st.text_input("Ghi chú / nguồn hệ số (bắt buộc nếu nhập hệ số trực tiếp)", key="th_gc")
            if st.button("💾 LƯU HÀNG & HỆ SỐ", type="primary", key="th_luu"):
                try:
                    ma = L.them_hang(store(), du_lieu_moi(), user(), so=so, ma_misa=ma_misa, ten_hang=ten, nhom=nhom,
                                     dvt_kho=dvt_kho, ghi_chu=ghi_chu, **kw)
                except L.LoiNghiepVu as e:
                    st.error(f"⚠ {e}")
                else:
                    du_lieu_moi()
                    st.success(f"Đã thêm hàng, mã quy cách **{ma}**. Kho dùng được ngay.")
        else:
            st.markdown('<div class="nho">Chọn mã MISA và điền ĐVT kho để chọn cách tính hệ số.</div>', unsafe_allow_html=True)

    with t2:
        h = _chon_hang_chung(data, "dh")
        if h is not None:
            qd = data["QUY_DOI"]
            qd = qd[qd["ma_hang"] == h["ma_hang"]]
            hs = L.he_so(data, h["ma_hang"], str(L.hom_nay()))
            if hs is None:
                st.error("Hàng này chưa có hệ số nào.")
            else:
                c = st.columns([5, 6], gap="medium")
                c[0].markdown(_thong_tin_he_so(hs))
                with c[0].expander(f"Lịch sử hệ số ({len(qd)} dòng)"):
                    st.dataframe(qd[["hieu_luc_tu", "he_so", "dvt_misa", "cach_tinh", "nguoi_duyet", "ngay_duyet"]],
                                 hide_index=True, width="stretch")
                with c[1]:
                    kw = _o_quy_doi(data, "dh", h["dvt_kho"], hs["dvt_misa"])
                    cc = st.columns([1.2, 3])
                    hl = cc[0].text_input("Hiệu lực từ ngày", value=f"{L.hom_nay():%d/%m/%Y}", key="dh_hl")
                    gc = cc[1].text_input("Ghi chú / nguồn (bắt buộc nếu nhập trực tiếp)", key="dh_gc")
                    st.caption("Phiếu đã lập giữ nguyên hệ số cũ; hệ số mới chỉ áp dụng cho phiếu lập sau.")
                    if st.button("💾 LƯU HỆ SỐ MỚI", type="primary", key="dh_luu"):
                        d = L.doc_ngay(hl)
                        if d is None:
                            st.error("Ngày hiệu lực chưa đọc được. Gõ dạng ngày/tháng/năm, ví dụ 27/11/2026.")
                        else:
                            try:
                                t = L.doi_he_so(store(), du_lieu_moi(), user(), h["ma_hang"], hieu_luc_tu=str(d), ghi_chu=gc, **kw)
                            except L.LoiNghiepVu as e:
                                st.error(f"⚠ {e}")
                            else:
                                du_lieu_moi()
                                st.success(f"Đã lưu: 1 {h['dvt_kho']} = {t['he_so']:g} {hs['dvt_misa']} từ {d:%d/%m/%Y}.")

    with t3:
        klr = L.klr_hien_hanh(data)
        da_sua = set(data["KHOI_LUONG_RIENG"]["vat_lieu"])
        c = st.columns([5, 6], gap="medium")
        c[0].dataframe(pd.DataFrame([dict(vat_lieu=k, klr=v, nguon="đã sửa trong app" if k in da_sua else "mặc định QT01")
                                     for k, v in klr.items()]).rename(columns={"vat_lieu": "Vật liệu", "klr": "g/cm³", "nguon": "Nguồn"}),
                       hide_index=True, width="stretch", height=330)
        with c[1]:
            st.markdown("**Đặt khối lượng riêng**")
            cc = st.columns([1.6, 1.2])
            vl = cc[0].selectbox("Vật liệu", list(klr) + ["➕ Vật liệu mới"], key="klr_vl")
            ten_vl = st.text_input("Tên vật liệu mới", key="klr_ten") if vl == "➕ Vật liệu mới" else vl
            gia_tri = cc[1].number_input("g/cm³", min_value=0.0, step=0.01, format="%.3f",
                                         value=float(klr.get(vl, 7.85)), key=f"klr_gt_{vl}")
            gc = st.text_input("Ghi chú / nguồn", key="klr_gc")
            st.caption("Đổi khối lượng riêng chỉ ảnh hưởng hàng thêm mới hoặc hệ số đổi SAU này. Hệ số đã lưu và phiếu đã lập không đổi.")
            if st.button("💾 LƯU KHỐI LƯỢNG RIÊNG", type="primary", key="klr_luu"):
                try:
                    L.dat_klr(store(), du_lieu_moi(), user(), ten_vl, gia_tri, gc)
                except L.LoiNghiepVu as e:
                    st.error(f"⚠ {e}")
                else:
                    du_lieu_moi()
                    st.success(f"Đã đặt {ten_vl} = {gia_tri:g} g/cm³.")
                    st.rerun()


# ---------------- Điều hướng ----------------

TRANG = {"chu": trang_chu, "ds_loai": trang_ds_loai, "phieu": trang_phieu, "nhan_ck": trang_nhan_ck, "ton": trang_ton,
         "ds_phieu": trang_ds_phieu, "bao_cao": trang_bao_cao, "kiem_ke": trang_kiem_ke,
         "khoa_so": trang_khoa_so, "danh_muc": trang_danh_muc, "nguoi_dung": trang_nguoi_dung,
         "quy_doi": trang_quy_doi, "dm_quan_tri": trang_dm_quan_tri}

if "user" not in st.session_state:
    trang_dang_nhap()
else:
    dau_trang()
    TRANG.get(st.session_state.get("trang", "chu"), trang_chu)()
