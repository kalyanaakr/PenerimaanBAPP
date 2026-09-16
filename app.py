"""
SISTEM PENERIMAAN BAPP
=======================
Aplikasi Streamlit untuk mencatat penerimaan BAPP fisik menggunakan
scanner barcode. Data master DAN hasil penerimaan sama-sama disimpan
di Google Spreadsheet (sheet "data").

Flow: Daftar Penerimaan BAPP -> (+) Buat Penerimaan Baru
(popup info) -> halaman scan BAPP -> Simpan -> Detail -> Print.

Cara menjalankan:
    streamlit run app.py

Lihat PANDUAN.md untuk instruksi instalasi & konfigurasi lengkap.
"""

import os
import io
import json
import time
import base64
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import sqlite3

import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from reportlab.pdfgen import canvas as pdfcanvas
from reportlab.graphics.barcode import code128
from reportlab.graphics.shapes import Drawing


# =====================================================================
# 1. KONFIGURASI
# =====================================================================

SPREADSHEET_ID = "1bgBsR4U5u5dgjTONE1RrMPHV2-prJXjfLd9Kf2Z5Jsc"
SHEET_NAME = "data"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CREDENTIALS_FILE = os.path.join(BASE_DIR, "credentials.json")
DB_FILE = os.path.join(BASE_DIR, "bapp_cadangan.db")

# Baris judul kolom di sheet "data" ada di baris ke-3 (baris 1-2 dipakai untuk
# info "Last Update" / "Summary"), data sebenarnya mulai baris ke-4.
BARIS_HEADER = 3

# Data master otomatis di-refresh ulang kalau sudah lebih tua dari ini (detik).
AUTO_REFRESH_DETIK = 5 * 60  # 5 menit

# Kolom data sekolah yang dibaca dari sheet "data" (sudah ada dari awal).
# "Nomor Urut Penerimaan" (tanpa akhiran) adalah nomor urut dari penerimaan
# PERTAMA -- dipakai di sini murni sebagai referensi tampilan, tidak diubah.
KOLOM_DATA_SEKOLAH = [
    "Nomor Transaksi",
    "NPSN",
    "Nama Sekolah",
    "Nomor Penerimaan",
    "Nomor Urut Penerimaan",
    "Serial Number",
    "Nama Koordinator",
    "Barcode Penerimaan",
    "Tanggal BAPP",
]

# Kolom untuk penerimaan KEDUA/BARU -- ini yang ditulis oleh aplikasi ini.
KOLOM_STATUS_BARU = "Status BAPP Fisik Kedua"
KOLOM_WAKTU_BARU = "Waktu BAPP diterima Baru"
KOLOM_NOMOR_BARU = "Nomor Penerimaan Baru"
KOLOM_PENGIRIM = "Nama Pengirim"
KOLOM_PIC = "PIC Penerimaan"
KOLOM_URUTAN_BARU = "Nomor Urut Penerimaan Baru"  # dipakai internal utk urutan simpan
KOLOM_URUTAN_PERTAMA = "Nomor Urut Penerimaan"    # referensi, sudah ada di sheet

KOLOM_TULIS = [KOLOM_STATUS_BARU, KOLOM_WAKTU_BARU, KOLOM_NOMOR_BARU, KOLOM_PENGIRIM, KOLOM_PIC, KOLOM_URUTAN_BARU]
KOLOM_WAJIB = KOLOM_DATA_SEKOLAH + KOLOM_TULIS

# Nilai status folder/penerimaan yang ditulis ke KOLOM_STATUS_BARU.
STATUS_OPEN = "OPEN"
STATUS_DITERIMA = "DITERIMA"

DAFTAR_DIREKTORAT_DEFAULT = ["SD", "SMP", "SMA", "SMK"]
UKURAN_HALAMAN_DAFTAR = 10
BAPP_PER_LEMBAR_PRINT = 30

# Ukuran kertas print Penerimaan BAPP: A4 portrait
KERTAS_PRINT = A4

st.set_page_config(page_title="Sistem Penerimaan BAPP", page_icon="📦", layout="wide")

st.markdown(
    """
    <style>
    .stButton > button { border-radius: 10px; font-weight: 500; transition: all 0.15s ease; }
    .stButton > button[kind="primary"] { background-color: #2563eb; border-color: #2563eb; }
    .stButton > button:hover { filter: brightness(0.95); border-color: #93c5fd; }
    div[data-testid="stTextInput"] input, div[data-testid="stSelectbox"] { border-radius: 8px; }
    div[data-testid="stHorizontalBlock"] {
        border-radius: 8px;
        transition: background-color 0.1s ease;
    }
    div[data-testid="stHorizontalBlock"]:hover {
        background-color: #f8fafc;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================================================================
# 2. DATABASE LOKAL -- HANYA untuk cadangan tersembunyi & pengaturan
# =====================================================================

def get_conn():
    return sqlite3.connect(DB_FILE, check_same_thread=False)


def init_db():
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("CREATE TABLE IF NOT EXISTS pengaturan (key TEXT PRIMARY KEY, value TEXT)")
    cur.execute("INSERT OR IGNORE INTO pengaturan (key, value) VALUES ('termin_penerimaan', '2')")
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS backup_penerimaan (
            nomor_penerimaan TEXT PRIMARY KEY,
            direktorat TEXT,
            waktu TEXT,
            jumlah_bapp INTEGER,
            detail_json TEXT,
            nama_pengirim TEXT,
            pic_penerimaan TEXT
        )
        """
    )
    for kolom in ("nama_pengirim", "pic_penerimaan"):
        try:
            cur.execute(f"ALTER TABLE backup_penerimaan ADD COLUMN {kolom} TEXT")
        except sqlite3.OperationalError:
            pass
    conn.commit()
    conn.close()


def get_setting(key, default=None):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT value FROM pengaturan WHERE key = ?", (key,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row else default


def set_setting(key, value):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO pengaturan (key, value) VALUES (?, ?) "
        "ON CONFLICT(key) DO UPDATE SET value = excluded.value",
        (key, value),
    )
    conn.commit()
    conn.close()


def simpan_cadangan_lokal(nomor_penerimaan, direktorat, waktu, scan_list, nama_pengirim="", pic_penerimaan=""):
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute(
            "INSERT OR REPLACE INTO backup_penerimaan "
            "(nomor_penerimaan, direktorat, waktu, jumlah_bapp, detail_json, nama_pengirim, pic_penerimaan) "
            "VALUES (?, ?, ?, ?, ?, ?, ?)",
            (nomor_penerimaan, direktorat, waktu, len(scan_list),
             json.dumps(scan_list, ensure_ascii=False), nama_pengirim, pic_penerimaan),
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


def hapus_cadangan_lokal(nomor_penerimaan):
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("DELETE FROM backup_penerimaan WHERE nomor_penerimaan = ?", (nomor_penerimaan,))
        conn.commit()
        conn.close()
    except Exception:
        pass


# =====================================================================
# 3. KONEKSI GOOGLE SPREADSHEET (baca + tulis)
# =====================================================================

def _panggil_dengan_retry(fungsi, percobaan=3, jeda_detik=2):
    """Menjalankan pemanggilan ke Google Sheets API dengan percobaan ulang
    otomatis. Gangguan koneksi sesaat (internet tidak stabil, firewall/
    antivirus/proxy) sering hilang sendiri kalau dicoba lagi beberapa detik
    kemudian, jadi tidak perlu langsung dianggap gagal total."""
    error_terakhir = None
    for percobaan_ke in range(percobaan):
        try:
            return fungsi()
        except Exception as e:
            error_terakhir = e
            if percobaan_ke < percobaan - 1:
                time.sleep(jeda_detik)
    raise error_terakhir


def _pesan_error_ramah(e):
    """Menerjemahkan error koneksi teknis jadi pesan yang lebih mudah
    dipahami, tanpa menyembunyikan pesan aslinya."""
    teks = str(e)
    penanda_jaringan = [
        "ConnectionReset", "Connection aborted", "Max retries exceeded",
        "RemoteDisconnected", "ConnectionError", "10054", "timed out", "Timeout",
    ]
    if any(p.lower() in teks.lower() for p in penanda_jaringan):
        return (
            f"Koneksi ke Google terputus sesaat (sudah dicoba ulang beberapa kali tapi "
            f"masih gagal). Biasanya karena internet yang kurang stabil, atau ada "
            f"firewall/antivirus/proxy yang mengganggu koneksi ke Google. Coba tekan "
            f"Refresh lagi, atau periksa koneksi internet Anda. Detail teknis: {teks}"
        )
    return teks


@st.cache_resource
def get_gsheet_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets"]
    try:
        ada_secret = "gcp_service_account" in st.secrets
    except Exception:
        ada_secret = False

    if ada_secret:
        # Dipakai saat aplikasi di-deploy online (mis. Streamlit Community
        # Cloud) -- kredensial diambil dari fitur Secrets, bukan file lokal.
        creds = Credentials.from_service_account_info(
            dict(st.secrets["gcp_service_account"]), scopes=scopes
        )
    else:
        # Dipakai saat dijalankan lokal di komputer sendiri.
        creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=scopes)
    return gspread.authorize(creds)


@st.cache_resource
def get_worksheet():
    client = get_gsheet_client()
    sh = _panggil_dengan_retry(lambda: client.open_by_key(SPREADSHEET_ID))
    return _panggil_dengan_retry(lambda: sh.worksheet(SHEET_NAME))


def load_master_data():
    ws = get_worksheet()
    semua = _panggil_dengan_retry(lambda: ws.get_all_values())
    if len(semua) < BARIS_HEADER:
        return pd.DataFrame(columns=["_baris_sheet"])

    header_mentah = [h.strip() for h in semua[BARIS_HEADER - 1]]
    n_kolom = len(header_mentah)

    header_bersih = []
    jumlah_pakai = {}
    for h in header_mentah:
        if h == "":
            header_bersih.append(None)
            continue
        jumlah_pakai[h] = jumlah_pakai.get(h, 0) + 1
        header_bersih.append(h if jumlah_pakai[h] == 1 else f"{h} ({jumlah_pakai[h]})")

    baris_data = [(row + [""] * n_kolom)[:n_kolom] for row in semua[BARIS_HEADER:]]

    df = pd.DataFrame(baris_data, columns=header_bersih)
    df = df.loc[:, [c for c in df.columns if c is not None]]
    df["_baris_sheet"] = range(BARIS_HEADER + 1, len(df) + BARIS_HEADER + 1)
    return df


def refresh_master_data():
    try:
        df = load_master_data()
        missing = [k for k in KOLOM_WAJIB if k not in df.columns]
        if missing:
            st.session_state.load_error = (
                f"Kolom berikut belum ada di sheet '{SHEET_NAME}': {', '.join(missing)}. "
                f"Tambahkan dulu di baris judul (baris {BARIS_HEADER}) Spreadsheet -- lihat PANDUAN.md."
            )
            return False
        st.session_state.master_df = df
        st.session_state.last_refresh = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        st.session_state.last_refresh_ts = datetime.now()
        st.session_state.load_error = None
        return True
    except FileNotFoundError:
        st.session_state.load_error = (
            f"File '{os.path.basename(CREDENTIALS_FILE)}' tidak ditemukan di folder aplikasi. "
            f"Ikuti PANDUAN.md bagian pembuatan credentials Google."
        )
        return False
    except Exception as e:
        st.session_state.load_error = _pesan_error_ramah(e)
        return False


def get_daftar_direktorat():
    df = st.session_state.get("master_df", pd.DataFrame())
    if not df.empty and "Direktorat" in df.columns:
        daftar = sorted(x for x in df["Direktorat"].dropna().unique().tolist() if str(x).strip())
        if daftar:
            return daftar
    return DAFTAR_DIREKTORAT_DEFAULT


# =====================================================================
# 4. NOMOR PENERIMAAN & PENYIMPANAN / PENGHAPUSAN DI SPREADSHEET
# =====================================================================

def generate_nomor_penerimaan(direktorat):
    termin = get_setting("termin_penerimaan", "2")
    prefix = f"{direktorat.upper()}{termin}-"

    df = st.session_state.get("master_df", pd.DataFrame())
    max_urut = 0
    if not df.empty and KOLOM_NOMOR_BARU in df.columns:
        existing = df[KOLOM_NOMOR_BARU].astype(str).str.strip()
        for val in existing[existing.str.startswith(prefix)]:
            suffix = val[len(prefix):]
            if suffix.isdigit():
                max_urut = max(max_urut, int(suffix))

    return f"{prefix}{max_urut + 1:03d}"


def _header_dan_idx_kolom_tulis(ws):
    header = [h.strip() for h in _panggil_dengan_retry(lambda: ws.row_values(BARIS_HEADER))]
    missing = [k for k in KOLOM_TULIS if k not in header]
    if missing:
        raise RuntimeError(f"Kolom berikut belum ada di sheet '{SHEET_NAME}': {', '.join(missing)}.")
    return {k: header.index(k) + 1 for k in KOLOM_TULIS}


def catat_scan_ke_spreadsheet(item, info, urutan):
    """Menulis SATU baris BAPP ke Spreadsheet segera setelah berhasil di-scan,
    berstatus OPEN -- supaya progress tidak hilang walau browser ditutup dan
    bisa dilanjutkan lagi nanti (lihat muat_penerimaan_open)."""
    ws = get_worksheet()
    idx = _header_dan_idx_kolom_tulis(ws)

    waktu_tulis = info["tanggal"].strftime("%d/%m/%Y")
    baris = item["_baris_sheet"]
    nilai = {
        KOLOM_STATUS_BARU: STATUS_OPEN,
        KOLOM_WAKTU_BARU: waktu_tulis,
        KOLOM_NOMOR_BARU: info["nomor_penerimaan"],
        KOLOM_PENGIRIM: info["nama_pengirim"],
        KOLOM_PIC: info["pic"],
        KOLOM_URUTAN_BARU: str(urutan),
    }
    updates = [{"range": rowcol_to_a1(baris, idx[k]), "values": [[v]]} for k, v in nilai.items()]
    _panggil_dengan_retry(lambda: ws.batch_update(updates, value_input_option="RAW"))

    df = st.session_state.master_df
    mask = df["_baris_sheet"] == baris
    for k, v in nilai.items():
        df.loc[mask, k] = v
    st.session_state.master_df = df

    # cadangan lokal di-update setiap scan supaya selalu mencerminkan progress terkini
    simpan_cadangan_lokal(
        info["nomor_penerimaan"], info["direktorat"], waktu_tulis,
        st.session_state.scan_list, info["nama_pengirim"], info["pic"],
    )


def hapus_satu_baris_bapp(item):
    """Mengosongkan kembali 6 kolom penerimaan baru untuk SATU baris BAPP saja
    (dipakai saat hapus 1 baris dari penerimaan yang masih OPEN)."""
    ws = get_worksheet()
    idx = _header_dan_idx_kolom_tulis(ws)
    baris = item["_baris_sheet"]
    updates = [{"range": rowcol_to_a1(baris, idx[k]), "values": [[""]]} for k in KOLOM_TULIS]
    _panggil_dengan_retry(lambda: ws.batch_update(updates, value_input_option="RAW"))

    df = st.session_state.master_df
    mask = df["_baris_sheet"] == baris
    for k in KOLOM_TULIS:
        df.loc[mask, k] = ""
    st.session_state.master_df = df


def tutup_penerimaan(nomor_penerimaan):
    """Menutup folder/penerimaan: ubah status semua baris terkait dari OPEN
    menjadi DITERIMA (final). Setelah ini datanya tidak diedit lagi oleh
    aplikasi (baik lewat scan/hapus baris)."""
    df = st.session_state.get("master_df", pd.DataFrame())
    baris_terkait = df[df[KOLOM_NOMOR_BARU].astype(str) == str(nomor_penerimaan)]
    if baris_terkait.empty:
        raise RuntimeError("Tidak ada BAPP dalam penerimaan ini untuk ditutup.")

    ws = get_worksheet()
    idx = _header_dan_idx_kolom_tulis(ws)

    updates = [
        {"range": rowcol_to_a1(int(item["_baris_sheet"]), idx[KOLOM_STATUS_BARU]), "values": [[STATUS_DITERIMA]]}
        for _, item in baris_terkait.iterrows()
    ]
    _panggil_dengan_retry(lambda: ws.batch_update(updates, value_input_option="RAW"))

    mask = df[KOLOM_NOMOR_BARU].astype(str) == str(nomor_penerimaan)
    df.loc[mask, KOLOM_STATUS_BARU] = STATUS_DITERIMA
    st.session_state.master_df = df


def muat_penerimaan_open(nomor_penerimaan):
    """Memuat ulang data penerimaan yang masih berstatus OPEN dari Spreadsheet
    ke session_state (penerimaan_aktif + scan_list), supaya operator bisa
    melanjutkan scan dari sesi/perangkat manapun. Mengembalikan (berhasil, pesan_error)."""
    df = st.session_state.get("master_df", pd.DataFrame())
    subset = df[df[KOLOM_NOMOR_BARU].astype(str) == str(nomor_penerimaan)].copy()
    if subset.empty:
        return False, "Data penerimaan tidak ditemukan."

    status_mentah = str(subset.iloc[0].get(KOLOM_STATUS_BARU, "")).strip().upper()
    if status_mentah != STATUS_OPEN:
        return False, "Penerimaan ini sudah DITERIMA (ditutup) dan tidak bisa diedit lagi."

    subset["_urut_num"] = pd.to_numeric(subset[KOLOM_URUTAN_BARU], errors="coerce")
    subset = subset.sort_values("_urut_num")

    scan_list_baru = []
    for _, r in subset.iterrows():
        scan_list_baru.append({
            "nomor_transaksi": str(r.get("Nomor Transaksi", "")).strip(),
            "npsn": str(r.get("NPSN", "")).strip(),
            "nama_sekolah": str(r.get("Nama Sekolah", "")).strip(),
            "nomor_penerimaan_pertama": str(r.get("Nomor Penerimaan", "")).strip(),
            "nomor_urut_pertama": str(r.get(KOLOM_URUTAN_PERTAMA, "")).strip(),
            "serial_number": str(r.get("Serial Number", "")).strip(),
            "nama_koordinator": str(r.get("Nama Koordinator", "")).strip(),
            "_baris_sheet": int(r["_baris_sheet"]),
        })

    baris_pertama = subset.iloc[0]
    waktu_str = str(baris_pertama.get(KOLOM_WAKTU_BARU, "")).strip()
    try:
        tanggal_obj = datetime.strptime(waktu_str.split(" ")[0], "%d/%m/%Y").date()
    except Exception:
        tanggal_obj = datetime.now().date()

    st.session_state.penerimaan_aktif = {
        "nomor_penerimaan": str(nomor_penerimaan),
        "direktorat": str(baris_pertama.get("Direktorat", "")),
        "nama_pengirim": str(baris_pertama.get(KOLOM_PENGIRIM, "")),
        "tanggal": tanggal_obj,
        "pic": str(baris_pertama.get(KOLOM_PIC, "")),
    }
    st.session_state.scan_list = scan_list_baru
    st.session_state.scan_message = None
    return True, None


def hapus_penerimaan(nomor_penerimaan):
    """Menghapus SELURUH penerimaan (biasanya untuk penerimaan yang masih
    OPEN dan ingin dibatalkan total): mengosongkan kembali 6 kolom penerimaan
    baru pada baris-baris terkait di Spreadsheet, sehingga BAPP tsb tersedia
    lagi untuk diterima ulang di penerimaan lain."""
    df = st.session_state.get("master_df", pd.DataFrame())
    if df.empty or KOLOM_NOMOR_BARU not in df.columns:
        return
    baris_terkait = df[df[KOLOM_NOMOR_BARU].astype(str) == str(nomor_penerimaan)]
    if baris_terkait.empty:
        return

    ws = get_worksheet()
    idx = _header_dan_idx_kolom_tulis(ws)

    updates = []
    for _, item in baris_terkait.iterrows():
        baris = int(item["_baris_sheet"])
        for kolom in KOLOM_TULIS:
            updates.append({"range": rowcol_to_a1(baris, idx[kolom]), "values": [[""]]})
    _panggil_dengan_retry(lambda: ws.batch_update(updates, value_input_option="RAW"))

    mask = df[KOLOM_NOMOR_BARU].astype(str) == str(nomor_penerimaan)
    for kolom in KOLOM_TULIS:
        df.loc[mask, kolom] = ""
    st.session_state.master_df = df

    hapus_cadangan_lokal(nomor_penerimaan)


def get_riwayat():
    kosong = pd.DataFrame(columns=["Nomor Penerimaan", "Tanggal", "Pengirim", "Direktorat", "Jumlah BAPP", "PIC", "Status"])
    df = st.session_state.get("master_df", pd.DataFrame())
    if df.empty or KOLOM_NOMOR_BARU not in df.columns:
        return kosong

    terisi = df[df[KOLOM_NOMOR_BARU].astype(str).str.strip() != ""]
    if terisi.empty:
        return kosong

    agg_dict = {
        "Direktorat": ("Direktorat", "first") if "Direktorat" in terisi.columns else (KOLOM_NOMOR_BARU, "first"),
        "Waktu": (KOLOM_WAKTU_BARU, "first"),
        "Jumlah BAPP": (KOLOM_NOMOR_BARU, "count"),
        "Status": (KOLOM_STATUS_BARU, "first"),
    }
    if KOLOM_PENGIRIM in terisi.columns:
        agg_dict["Pengirim"] = (KOLOM_PENGIRIM, "first")
    if KOLOM_PIC in terisi.columns:
        agg_dict["PIC"] = (KOLOM_PIC, "first")

    ringkasan = (
        terisi.groupby(KOLOM_NOMOR_BARU).agg(**agg_dict).reset_index()
        .rename(columns={KOLOM_NOMOR_BARU: "Nomor Penerimaan"})
    )
    # nilai lama sebelum fitur status ada (mis. "Diterima") dianggap tetap DITERIMA
    ringkasan["Status"] = ringkasan["Status"].apply(
        lambda v: STATUS_OPEN if str(v).strip().upper() == STATUS_OPEN else STATUS_DITERIMA
    )
    # ambil bagian tanggal saja -- kompatibel dengan data lama yang masih
    # menyertakan jam ("03/09/2026 14:30") maupun data baru tanpa jam ("03/09/2026")
    tanggal_saja = ringkasan["Waktu"].astype(str).str.split(" ").str[0]
    ringkasan["_waktu_dt"] = pd.to_datetime(tanggal_saja, format="%d/%m/%Y", errors="coerce")
    ringkasan["Tanggal"] = ringkasan["_waktu_dt"].dt.strftime("%d/%m/%Y")
    ringkasan = (
        ringkasan.sort_values(
            ["_waktu_dt", "Nomor Penerimaan"],
            ascending=[False, False],
            na_position="last",
        )
        .drop(columns=["_waktu_dt", "Waktu"])
    )

    for kolom in ["Pengirim", "PIC"]:
        if kolom not in ringkasan.columns:
            ringkasan[kolom] = ""

    return ringkasan.reset_index(drop=True)


def ekstrak_nomor_penerimaan_pertama(nilai):
    """Nomor Penerimaan Pertama dari spreadsheet formatnya 'BAPP-RCV-SD-0006' --
    tampilkan cuma bagian setelah 'BAPP-RCV-' (mis. 'SD-0006')."""
    nilai = str(nilai).strip()
    prefix = "BAPP-RCV-"
    if nilai.upper().startswith(prefix):
        return nilai[len(prefix):]
    return nilai


def format_tanggal_bapp(nilai):
    """Format kolom 'Tanggal BAPP' jadi DD/MM/YYYY. Kalau sumbernya masih
    mengandung jam, bagian jam dibuang. Kalau formatnya tidak dikenali,
    tampilkan apa adanya (bukan error)."""
    teks = str(nilai).strip()
    if not teks:
        return "-"
    teks_tanggal = teks.split(" ")[0]
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%m/%d/%Y", "%d-%m-%Y"):
        try:
            return datetime.strptime(teks_tanggal, fmt).strftime("%d/%m/%Y")
        except ValueError:
            continue
    return teks


def get_detail_penerimaan(nomor_penerimaan):
    df = st.session_state.get("master_df", pd.DataFrame())
    if df.empty or KOLOM_NOMOR_BARU not in df.columns:
        return pd.DataFrame(), {}

    subset = df[df[KOLOM_NOMOR_BARU].astype(str) == str(nomor_penerimaan)].copy()
    if subset.empty:
        return pd.DataFrame(), {}

    if KOLOM_URUTAN_BARU in subset.columns:
        subset["_urut_num"] = pd.to_numeric(subset[KOLOM_URUTAN_BARU], errors="coerce")
        subset = subset.sort_values("_urut_num")

    baris_pertama = subset.iloc[0]
    status_mentah = str(baris_pertama.get(KOLOM_STATUS_BARU, "")).strip().upper()
    info = {
        "nomor_penerimaan": str(nomor_penerimaan),
        "direktorat": str(baris_pertama.get("Direktorat", "")),
        "pengirim": str(baris_pertama.get(KOLOM_PENGIRIM, "")),
        "pic": str(baris_pertama.get(KOLOM_PIC, "")),
        "waktu": str(baris_pertama.get(KOLOM_WAKTU_BARU, "")),
        "jumlah": len(subset),
        "status": STATUS_OPEN if status_mentah == STATUS_OPEN else STATUS_DITERIMA,
    }

    def kolom_atau_kosong(nama):
        return subset[nama] if nama in subset.columns else [""] * len(subset)

    tabel = pd.DataFrame({
        "Nomor": range(1, len(subset) + 1),
        "Nomor Transaksi": kolom_atau_kosong("Nomor Transaksi"),
        "NPSN": kolom_atau_kosong("NPSN"),
        "Nama Sekolah": kolom_atau_kosong("Nama Sekolah"),
        "Tanggal BAPP": [format_tanggal_bapp(v) for v in kolom_atau_kosong("Tanggal BAPP")],
        "Nomor Penerimaan Pertama": [ekstrak_nomor_penerimaan_pertama(v) for v in kolom_atau_kosong("Nomor Penerimaan")],
        "Barcode Penerimaan": kolom_atau_kosong("Barcode Penerimaan"),
        "Nomor Urut": kolom_atau_kosong(KOLOM_URUTAN_PERTAMA),
        "Serial Number": kolom_atau_kosong("Serial Number"),
        "Nama Koordinator": kolom_atau_kosong("Nama Koordinator"),
    }).reset_index(drop=True)

    return tabel, info


# =====================================================================
# 5. LOGIKA SCAN
# =====================================================================

def proses_scan(nomor_transaksi):
    df = st.session_state.get("master_df", pd.DataFrame())
    if df.empty or "Nomor Transaksi" not in df.columns:
        st.session_state.scan_message = (
            "error", "Data master belum tersedia. Buka menu Pengaturan lalu tekan Refresh Data.",
        )
        return

    nomor_transaksi = nomor_transaksi.strip()
    cocok = df[df["Nomor Transaksi"].astype(str).str.strip() == nomor_transaksi]

    if cocok.empty:
        st.session_state.scan_message = ("error", f"❌ Nomor transaksi tidak ditemukan: {nomor_transaksi}")
        return

    baris = cocok.iloc[0]
    nama_sekolah = str(baris.get("Nama Sekolah", "")).strip()

    # Batasi per-direktorat: BAPP harus punya Direktorat yang sama dengan
    # penerimaan yang sedang dibuat -- kalau beda, tolak.
    info_aktif = st.session_state.get("penerimaan_aktif") or {}
    direktorat_aktif = str(info_aktif.get("direktorat", "")).strip().upper()
    direktorat_bapp = str(baris.get("Direktorat", "")).strip().upper()
    if direktorat_aktif and direktorat_bapp and direktorat_aktif != direktorat_bapp:
        st.session_state.scan_message = (
            "error",
            f"❌ BAPP ini milik Direktorat {direktorat_bapp}, bukan {direktorat_aktif} "
            f"(Direktorat penerimaan yang sedang berjalan). Tidak bisa ditambahkan.",
        )
        return

    nomor_lama = str(baris.get(KOLOM_NOMOR_BARU, "")).strip()
    if nomor_lama:
        st.session_state.scan_message = (
            "warning",
            f"⚠️ BAPP sudah diterima — {nomor_transaksi} sudah tercatat pada penerimaan {nomor_lama}.",
        )
        return

    serial_number = str(baris.get("Serial Number", "")).strip()

    for item in st.session_state.scan_list:
        if item["nomor_transaksi"] == nomor_transaksi:
            st.session_state.scan_message = ("warning", "⚠️ BAPP sudah ada dalam daftar penerimaan ini.")
            return
        if serial_number and item["serial_number"] == serial_number:
            st.session_state.scan_message = (
                "warning", f"⚠️ Serial Number sudah ada dalam daftar penerimaan ini: {serial_number}",
            )
            return

    st.session_state.scan_list.append(
        {
            "nomor_transaksi": nomor_transaksi,
            "npsn": str(baris.get("NPSN", "")).strip(),
            "nama_sekolah": nama_sekolah,
            "nomor_penerimaan_pertama": str(baris.get("Nomor Penerimaan", "")).strip(),
            "nomor_urut_pertama": str(baris.get(KOLOM_URUTAN_PERTAMA, "")).strip(),
            "serial_number": serial_number,
            "nama_koordinator": str(baris.get("Nama Koordinator", "")).strip(),
            "_baris_sheet": int(baris["_baris_sheet"]),
        }
    )
    st.session_state.scan_message = ("success", f"✅ BAPP berhasil ditambahkan — {nomor_transaksi} ({nama_sekolah})")


def handle_scan_input():
    nomor = st.session_state.input_scan.strip()
    st.session_state.input_scan = ""
    if not nomor:
        return

    jumlah_sebelum = len(st.session_state.scan_list)
    proses_scan(nomor)
    if len(st.session_state.scan_list) <= jumlah_sebelum:
        return  # ditolak (tidak ditemukan/duplikat) -- tidak ada yg perlu ditulis

    # berhasil masuk ke list lokal -> langsung tulis ke Spreadsheet sbg OPEN
    info = st.session_state.get("penerimaan_aktif")
    urutan_baru = len(st.session_state.scan_list)
    item_baru = st.session_state.scan_list[-1]
    try:
        with st.spinner("Menyimpan ke Spreadsheet..."):
            catat_scan_ke_spreadsheet(item_baru, info, urutan_baru)
    except Exception as e:
        # gagal tersimpan -> batalkan penambahan di list lokal juga, supaya
        # yang tampil di layar selalu sama dengan yang benar-benar tersimpan
        st.session_state.scan_list.pop()
        st.session_state.scan_message = (
            "error",
            f"BAPP gagal tersimpan ke Spreadsheet: {_pesan_error_ramah(e)} Silakan scan ulang.",
        )


def autofocus_scan_input():
    components.html(
        """
        <script>
        setTimeout(function() {
            const doc = window.parent.document;
            const input = doc.querySelector('input[aria-label="scan_nomor_transaksi"]');
            if (input) { input.focus(); }
        }, 150);
        </script>
        """,
        height=0,
    )


# =====================================================================
# 6. CETAK PDF PENERIMAAN
# =====================================================================

class NumberedCanvas(pdfcanvas.Canvas):
    def __init__(self, *args, **kwargs):
        pdfcanvas.Canvas.__init__(self, *args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        total_halaman = len(self._saved_page_states)

        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._gambar_nomor_halaman(total_halaman)
            pdfcanvas.Canvas.showPage(self)

        pdfcanvas.Canvas.save(self)

    def _gambar_nomor_halaman(self, total_halaman):
        self.setFont("Helvetica", 8)

        lebar_halaman = self._pagesize[0]

        self.drawRightString(
            lebar_halaman - 1 * cm,
            1.1 * cm,
            f"Halaman {self._pageNumber} dari {total_halaman}"
        )


# =====================================================================
# FORMAT ISI CELL TABEL
# =====================================================================

NAMA_HARI_ID = {
    0: "Senin", 1: "Selasa", 2: "Rabu", 3: "Kamis",
    4: "Jumat", 5: "Sabtu", 6: "Minggu",
}


def tambah_nama_hari(tanggal_str):
    """Menambahkan nama hari Indonesia di depan tanggal format DD/MM/YYYY,
    mis. '15/09/2026' -> 'Selasa, 15/09/2026'. Kalau formatnya tidak
    dikenali, kembalikan tanggal apa adanya (bukan error)."""
    try:
        tgl = datetime.strptime(str(tanggal_str).strip(), "%d/%m/%Y")
        return f"{NAMA_HARI_ID[tgl.weekday()]}, {tanggal_str}"
    except Exception:
        return tanggal_str


def buat_sel_barcode(nilai, style_teks=None):
    """
    Barcode Penerimaan ditampilkan sebagai TEKS biasa,
    bukan barcode gambar.
    """

    teks = str(nilai).strip()

    if not teks:
        teks = "-"

    if style_teks:
        return Paragraph(teks, style_teks)

    return teks


# =====================================================================
# CETAK PDF PENERIMAAN
# =====================================================================

def buat_pdf_penerimaan(info, tabel_df):
    """
    Membuat PDF Bukti Penerimaan BAPP.

    Layout:
    - Portrait A4 (21 x 29,7 cm)
    - Maksimal 40 BAPP per lembar
    - Layout dibuat lebih lega seperti format referensi
    - Tidak menggunakan PageBreak yang dapat menyebabkan halaman kosong
    - Tanda tangan hanya di halaman terakhir
    - Ringkasan bundle (Nomor Penerimaan Pertama) hanya di halaman terakhir,
      dihitung dari SELURUH BAPP di penerimaan ini -- supaya Tim Sortir tahu
      bundle apa saja yang perlu dicari.
    """
    buffer = io.BytesIO()

    # ================================================================
    # UKURAN KERTAS
    # ================================================================
    PAGE_W, PAGE_H = KERTAS_PRINT

    # Margin dibuat cukup lega agar hasil mirip format referensi.
    margin_left = 1.25 * cm
    margin_right = 1.25 * cm
    margin_top = 0.85 * cm
    margin_bottom = 0.9 * cm

    doc = SimpleDocTemplate(
        buffer,
        pagesize=KERTAS_PRINT,
        topMargin=margin_top,
        bottomMargin=margin_bottom,
        leftMargin=margin_left,
        rightMargin=margin_right,
        title="BUKTI PENERIMAAN BAPP TAHAP 2",
        author="PT. PYX SOLUSI TEKNOLOGI",
    )

    styles = getSampleStyleSheet()

    # ================================================================
    # STYLE
    # ================================================================
    judul_style = ParagraphStyle(
        "JudulBAPP",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=15,
        spaceBefore=0,
        spaceAfter=6,
        alignment=1,
    )

    sel_style = ParagraphStyle(
        "SelBAPP",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=7.5,
        leading=7.8,
        alignment=1,
        spaceBefore=0,
        spaceAfter=0,
    )

    header_sel_style = ParagraphStyle(
        "HeaderBAPP",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=8,
        leading=8.2,
        alignment=1,
        spaceBefore=0,
        spaceAfter=0,
    )

    info_label_style = ParagraphStyle(
        "InfoLabelBAPP",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=10.5,
        alignment=0,
        spaceBefore=0,
        spaceAfter=0,
    )

    info_value_style = ParagraphStyle(
        "InfoValueBAPP",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10,
        leading=10.5,
        alignment=0,
        spaceBefore=0,
        spaceAfter=0,
    )

    judul_ringkasan_style = ParagraphStyle(
        "JudulRingkasanBAPP",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=10.5,
        leading=11.5,
        spaceBefore=0,
        spaceAfter=4,
        alignment=1,
    )

    def sel(txt):
        nilai = str(txt).strip()
        if not nilai:
            nilai = "-"
        return Paragraph(nilai, sel_style)

    def header_sel(txt):
        return Paragraph(str(txt), header_sel_style)

    def info_label(txt):
        return Paragraph(str(txt), info_label_style)

    def info_value(txt):
        nilai = str(txt).strip() if txt is not None else "-"
        if not nilai:
            nilai = "-"
        return Paragraph(nilai, info_value_style)

    # ================================================================
    # TANGGAL
    # ================================================================
    tanggal_saja = (
        str(info.get("waktu", "")).split(" ")[0]
        if info.get("waktu")
        else "-"
    )

    # ================================================================
    # LEBAR TABEL
    #
    # Total = 21.5 cm.
    # Kolom Nomor Urut & Nama Koordinator dihapus (tidak dicetak lagi),
    # lebarnya dibagi ulang ke kolom lain -- termasuk Tanggal BAPP yang
    # sekarang lebih lebar karena menampilkan nama hari juga.
    # ================================================================
    lebar_kolom = [
        0.65 * cm,   # No
        2.45 * cm,   # Nomor Transaksi
        1.35 * cm,   # NPSN
        3.55 * cm,   # Nama Sekolah
        2.55 * cm,   # Tanggal BAPP (+ nama hari)
        1.85 * cm,   # Barcode Penerimaan
        1.75 * cm,   # Nomor Penerimaan 1
        4.35 * cm,   # Serial Number
    ]

    # ================================================================
    # RINGKASAN BUNDLE ASAL (dihitung SEKALI dari seluruh tabel_df,
    # dipakai nanti di halaman terakhir saja)
    # ================================================================
    if "Nomor Penerimaan Pertama" in tabel_df.columns:
        ringkasan_bundle = (
            tabel_df["Nomor Penerimaan Pertama"]
            .value_counts()
            .reset_index()
        )
        ringkasan_bundle.columns = ["Nomor Penerimaan Pertama", "Jumlah BAPP"]
        ringkasan_bundle = ringkasan_bundle.sort_values("Nomor Penerimaan Pertama")
    else:
        ringkasan_bundle = pd.DataFrame(columns=["Nomor Penerimaan Pertama", "Jumlah BAPP"])

    # ================================================================
    # DATA / JUMLAH HALAMAN
    # ================================================================
    total_baris = len(tabel_df)

    total_lembar = max(
        1,
        -(-total_baris // BAPP_PER_LEMBAR_PRINT)
    )

    flow = []

    # ================================================================
    # LOOP HALAMAN
    # ================================================================
    for lembar in range(total_lembar):

        mulai = lembar * BAPP_PER_LEMBAR_PRINT
        selesai = (lembar + 1) * BAPP_PER_LEMBAR_PRINT
        potongan = tabel_df.iloc[mulai:selesai]

        # ------------------------------------------------------------
        # JUDUL
        # ------------------------------------------------------------
        flow.append(
            Paragraph(
                "BUKTI PENERIMAAN BAPP TAHAP 2",
                judul_style
            )
        )

        # ------------------------------------------------------------
        # INFO PENERIMAAN
        # ------------------------------------------------------------
        info_rows = [
            [
                info_label("Nomor Penerimaan"),
                info_value(":"),
                info_value(info.get("nomor_penerimaan") or "-"),
                info_label("Tanggal"),
                info_value(":"),
                info_value(tanggal_saja),
            ],
            [
                info_label("Pengirim"),
                info_value(":"),
                info_value(info.get("pengirim") or "-"),
                info_label("PIC Penerimaan"),
                info_value(":"),
                info_value(info.get("pic") or "-"),
            ],
            [
                info_label("Direktorat"),
                info_value(":"),
                info_value(info.get("direktorat") or "-"),
                info_label("Jumlah BAPP"),
                info_value(":"),
                info_value(str(info.get("jumlah", ""))),
            ],
        ]

        t_info = Table(
            info_rows,
            colWidths=[
                2.35 * cm,
                0.25 * cm,
                6.35 * cm,
                2.35 * cm,
                0.25 * cm,
                6.35 * cm,
            ],
            rowHeights=[
                0.62 * cm,
                0.62 * cm,
                0.62 * cm,
            ],
        )

        t_info.setStyle(
            TableStyle([
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 0), (1, -1), "CENTER"),
                ("ALIGN", (4, 0), (4, -1), "CENTER"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 1),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ])
        )

        flow.append(t_info)

        # Jarak pendek sebelum tabel.
        flow.append(Spacer(1, 0.42 * cm))

        # ------------------------------------------------------------
        # HEADER TABEL
        # ------------------------------------------------------------
        header_tabel = [
            header_sel(t)
            for t in [
                "No",
                "Nomor Transaksi",
                "NPSN",
                "Nama Sekolah",
                "Tanggal BAPP",
                "Barcode<br/>Penerimaan",
                "Nomor<br/>Penerimaan 1",
                "Serial Number",
            ]
        ]

        data_tabel = [header_tabel]

        # ------------------------------------------------------------
        # ISI TABEL
        # ------------------------------------------------------------
        for _, r in potongan.iterrows():
            data_tabel.append([
                sel(r["Nomor"]),
                sel(r["Nomor Transaksi"]),
                sel(r["NPSN"]),
                sel(r["Nama Sekolah"]),
                sel(tambah_nama_hari(r["Tanggal BAPP"])),
                buat_sel_barcode(
                    r["Barcode Penerimaan"],
                    style_teks=sel_style
                ),
                sel(r["Nomor Penerimaan Pertama"]),
                sel(r["Serial Number"]),
            ])

        # ------------------------------------------------------------
        # TABEL
        #
        # Tinggi 0.55 cm x 40 = 22 cm (disesuaikan naik karena font
        # tabel diperbesar dari 5.6pt ke 7.5pt).
        # ------------------------------------------------------------
        row_heights = [0.85 * cm] + [
            0.55 * cm for _ in range(len(potongan))
        ]

        t = Table(
            data_tabel,
            colWidths=lebar_kolom,
            rowHeights=row_heights,
            repeatRows=1,
            splitByRow=1,
            hAlign="CENTER",
        )

        t.setStyle(
            TableStyle([
                # Header
                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#e5e7eb")
                ),
                (
                    "TEXTCOLOR",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#111827")
                ),

                # Grid
                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.4,
                    colors.HexColor("#111827")
                ),

                # Font
                (
                    "FONTNAME",
                    (0, 1),
                    (-1, -1),
                    "Helvetica"
                ),
                (
                    "FONTSIZE",
                    (0, 1),
                    (-1, -1),
                    7.5
                ),

                # Padding
                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),
                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    0
                ),
                (
                    "LEFTPADDING",
                    (0, 0),
                    (-1, -1),
                    1
                ),
                (
                    "RIGHTPADDING",
                    (0, 0),
                    (-1, -1),
                    1
                ),

                # Alignment
                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER"
                ),
                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),
            ])
        )

        flow.append(t)

        # ------------------------------------------------------------
        # TANDA TANGAN - HANYA HALAMAN TERAKHIR
        # ------------------------------------------------------------
        if lembar == total_lembar - 1:

            flow.append(Spacer(1, 1.0 * cm))

            pengirim_label = (
                info.get("pengirim")
                or "..........................."
            )

            pic_label = (
                info.get("pic")
                or "..........................."
            )

            kolom_kosong = "(" + " " * 18 + ")"

            ttd_rows = [
                [
                    "Pengirim",
                    "",
                    "Tim Sortir",
                    "",
                    "Tim Scan",
                    "",
                    "PIC Penerimaan"
                ],
                [
                    "", "", "", "", "", "", ""
                ],
                [
                    "", "", "", "", "", "", ""
                ],
                [
                    f"({pengirim_label})",
                    "",
                    kolom_kosong,
                    "",
                    kolom_kosong,
                    "",
                    f"({pic_label})"
                ],
            ]

            t_ttd = Table(
                ttd_rows,
                colWidths=[
                    3.95 * cm,
                    0.65 * cm,
                    3.95 * cm,
                    0.65 * cm,
                    3.95 * cm,
                    0.65 * cm,
                    3.95 * cm,
                ],
                rowHeights=[
                    0.55 * cm,
                    1.8 * cm,
                    0.45 * cm,
                    0.55 * cm,
                ],
                hAlign="CENTER",
            )

            t_ttd.setStyle(
                TableStyle([
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                    ("FONTNAME", (2, 0), (2, 0), "Helvetica-Bold"),
                    ("FONTNAME", (4, 0), (4, 0), "Helvetica-Bold"),
                    ("FONTNAME", (6, 0), (6, 0), "Helvetica-Bold"),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 0),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                    ("TOPPADDING", (0, 0), (-1, -1), 0),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ])
            )

            flow.append(t_ttd)
            flow.append(Spacer(1, 0.35 * cm))

        # ------------------------------------------------------------
        # RINGKASAN BUNDLE ASAL (Nomor Penerimaan Pertama) -- HANYA
        # HALAMAN TERAKHIR. Dihitung dari SELURUH BAPP di penerimaan
        # ini (bukan cuma yang ada di halaman terakhir), supaya Tim
        # Sortir tahu bundle apa saja yang perlu dicari & berapa
        # banyak BAPP dari masing-masing bundle itu.
        # ------------------------------------------------------------
        if lembar == total_lembar - 1 and not ringkasan_bundle.empty:

            flow.append(PageBreak())
            flow.append(Spacer(1, 0.35 * cm))
            flow.append(
                Paragraph(
                    "Ringkasan Bundle Asal (untuk Tim Sortir)",
                    judul_ringkasan_style
                )
            )
            flow.append(Spacer(1, 0.45 * cm))

            # Informasi penerimaan pada halaman ringkasan bundle
            info_ringkasan_rows = [
                [info_label("Nomor Penerimaan"), info_value(info.get("nomor_penerimaan") or "-"), info_label("Tanggal"), info_value(tanggal_saja)],
                [info_label("Pengirim"), info_value(info.get("pengirim") or "-"), info_label("PIC Penerimaan"), info_value(info.get("pic") or "-")],
                [info_label("Direktorat"), info_value(info.get("direktorat") or "-"), info_label("Jumlah BAPP"), info_value(str(info.get("jumlah", "")))],
            ]
            flow.append(
                Table(
                    info_ringkasan_rows,
                    colWidths=[3.0 * cm, 6.0 * cm, 3.0 * cm, 6.0 * cm],
                    rowHeights=[0.85 * cm, 0.85 * cm, 0.85 * cm],
                    hAlign="CENTER",
                    style=TableStyle([
                        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                        ("LEFTPADDING", (0, 0), (-1, -1), 4),
                        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                        ("TOPPADDING", (0, 0), (-1, -1), 5),
                        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
                    ])
                )
            )
            flow.append(Spacer(1, 0.55 * cm))

            ringkasan_rows = [["Nomor Penerimaan Pertama", "Jumlah BAPP"]]
            for _, r in ringkasan_bundle.iterrows():
                ringkasan_rows.append([
                    str(r["Nomor Penerimaan Pertama"]),
                    str(r["Jumlah BAPP"]),
                ])

            t_ringkasan = Table(
                ringkasan_rows,
                colWidths=[6 * cm, 3 * cm],
                hAlign="CENTER",
            )

            t_ringkasan.setStyle(
                TableStyle([
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e5e7eb")),
                    ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                    ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 9.5),
                    ("LEFTPADDING", (0, 0), (-1, -1), 4),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ])
            )

            flow.append(t_ringkasan)


        # ------------------------------------------------------------
        # PAGE BREAK
        #
        # PageBreak hanya diberikan kalau memang masih ada halaman
        # berikutnya. Karena tabel sekarang dibuat cukup pendek,
        # tidak akan terjadi split tabel -> PageBreak ganda.
        # ------------------------------------------------------------
        if lembar < total_lembar - 1:
            flow.append(PageBreak())

    # ================================================================
    # BUILD
    # ================================================================
    doc.build(
        flow,
        canvasmaker=NumberedCanvas
    )

    buffer.seek(0)
    return buffer.getvalue()


# =====================================================================
# 7. KOMPONEN TAMPILAN KECIL (card, badge)
# =====================================================================

def render_badge(teks, warna="abu"):
    palet = {
        "hijau": ("#166534", "#dcfce7"),
        "kuning": ("#92400e", "#fef3c7"),
        "abu": ("#374151", "#f3f4f6"),
        "biru": ("#1d4ed8", "#dbeafe"),
    }
    fg, bg = palet.get(warna, palet["abu"])
    return (
        f'<span style="background:{bg};color:{fg};padding:3px 10px;border-radius:999px;'
        f'font-size:12px;font-weight:600;white-space:nowrap;">{teks}</span>'
    )


def render_tombol_pdf(pdf_bytes, nama_file):
    """Menampilkan 2 tombol berdampingan (Preview & Print + Unduh) dalam SATU
    blok HTML yang sama, supaya ukuran & posisinya simetris persis -- kalau
    salah satu pakai tombol Streamlit asli dan satunya custom HTML, tingginya
    suka beda sedikit dan jadi tidak sejajar.

    Keduanya pakai teknik Blob URL lewat JavaScript (bukan link data: biasa) --
    browser modern seperti Chrome/Edge tidak lagi membuka PDF langsung dari
    link data:application/pdf;base64,... (yang muncul cuma teks base64 mentah).
    Blob URL didukung penuh baik untuk dibuka di viewer PDF maupun diunduh."""
    b64 = base64.b64encode(pdf_bytes).decode("utf-8")
    components.html(
        f"""
        <div style="display:flex; gap:0.75rem; width:100%; font-family:inherit;">
            <button id="btnPreviewPdf" style="
                flex:1; height:2.6rem; border-radius:10px; border:1px solid #d1d5db;
                background-color:#ffffff; color:#111827; font-weight:500; cursor:pointer;
                font-family:inherit; font-size:0.95rem;
            ">👁️ Preview &amp; Print PDF</button>
            <button id="btnUnduhPdf" style="
                flex:1; height:2.6rem; border-radius:10px; border:1px solid #2563eb;
                background-color:#2563eb; color:#ffffff; font-weight:500; cursor:pointer;
                font-family:inherit; font-size:0.95rem;
            ">⬇️ Unduh PDF</button>
        </div>
        <script>
        const base64Data = "{b64}";
        function buatBlobPdf() {{
            const byteChars = atob(base64Data);
            const byteNumbers = new Array(byteChars.length);
            for (let i = 0; i < byteChars.length; i++) {{
                byteNumbers[i] = byteChars.charCodeAt(i);
            }}
            const byteArray = new Uint8Array(byteNumbers);
            return new Blob([byteArray], {{ type: "application/pdf" }});
        }}
        document.getElementById("btnPreviewPdf").addEventListener("click", function() {{
            const blobUrl = URL.createObjectURL(buatBlobPdf());
            window.open(blobUrl, "_blank");
        }});
        document.getElementById("btnUnduhPdf").addEventListener("click", function() {{
            const blobUrl = URL.createObjectURL(buatBlobPdf());
            const a = document.createElement("a");
            a.href = blobUrl;
            a.download = "{nama_file}";
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
        }});
        </script>
        """,
        height=55,
    )


NAMA_BULAN_ID = {
    1: "Januari", 2: "Februari", 3: "Maret", 4: "April", 5: "Mei", 6: "Juni",
    7: "Juli", 8: "Agustus", 9: "September", 10: "Oktober", 11: "November", 12: "Desember",
}


def format_tanggal_indonesia(tanggal):
    """Format tanggal manual ke Bahasa Indonesia (mis. '03 September 2026'),
    tidak bergantung pada locale sistem operasi (yang di banyak komputer
    Windows default-nya Inggris, jadi %B akan salah menampilkan nama bulan)."""
    return f"{tanggal.day:02d} {NAMA_BULAN_ID[tanggal.month]} {tanggal.year}"


# =====================================================================
# 8. POPUP / DIALOG
# =====================================================================

@st.dialog("BUAT PENERIMAAN BARU")
def dialog_buat_penerimaan():
    # Placeholder di posisi paling atas (sesuai mockup), isinya baru ditentukan
    # setelah Direktorat diketahui di bawah -- st.empty() memungkinkan ini.
    nomor_placeholder = st.empty()

    nama_pengirim = st.text_input("Nama Pengirim", key="dlg_pengirim")

    opsi_direktorat = get_daftar_direktorat() + ["Lainnya (isi manual)"]
    pilihan_direktorat = st.selectbox("Direktorat", opsi_direktorat, key="dlg_direktorat_pilihan")
    if pilihan_direktorat == "Lainnya (isi manual)":
        direktorat = st.text_input("Ketik kode Direktorat", key="dlg_direktorat_manual").strip().upper()
    else:
        direktorat = pilihan_direktorat

    tanggal_penerimaan = st.date_input("Tanggal Penerimaan", value=datetime.now().date(), key="dlg_tanggal")
    pic_penerimaan = st.text_input("PIC Penerimaan", key="dlg_pic")

    nomor_preview = generate_nomor_penerimaan(direktorat) if direktorat else "-"
    # Sengaja TANPA key= -- field ini cuma tampilan (disabled), nilainya harus
    # ikut berubah setiap kali Direktorat diganti. Kalau pakai key, Streamlit
    # akan mempertahankan nilai lama dari session_state dan mengabaikan
    # parameter value= yang baru (ini penyebab bug "nomor tidak ikut berubah").
    nomor_placeholder.text_input("Nomor Penerimaan 🔒", value=nomor_preview, disabled=True)

    st.markdown("")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("BATAL", use_container_width=True):
            st.rerun()
    with c2:
        siap = bool(direktorat and nama_pengirim.strip() and pic_penerimaan.strip())
        if st.button("LANJUT", type="primary", use_container_width=True, disabled=not siap):
            st.session_state.penerimaan_aktif = {
                "nomor_penerimaan": nomor_preview,
                "direktorat": direktorat,
                "nama_pengirim": nama_pengirim.strip(),
                "tanggal": tanggal_penerimaan,
                "pic": pic_penerimaan.strip(),
            }
            st.session_state.scan_list = []
            st.session_state.scan_message = None
            st.session_state.halaman = "form_baru"
            st.rerun()


def buka_dialog_penerimaan_baru():
    """Bersihkan field popup sebelum dibuka, supaya tidak ada data yang
    'nyangkut' dari percobaan sebelumnya (mis. setelah klik BATAL). Kalau ada
    penerimaan OPEN yang sedang aktif dikerjakan di sesi ini, arahkan ke sana
    dulu (datanya sendiri sudah aman tersimpan di Spreadsheet, tapi baiknya
    diselesaikan/ditutup dulu sebelum mulai penerimaan lain)."""
    if st.session_state.get("scan_list") and st.session_state.get("penerimaan_aktif"):
        st.session_state.halaman = "form_baru"
        st.toast("Ada penerimaan OPEN yang sedang dikerjakan -- selesaikan/tutup dulu.", icon="⚠️")
        st.rerun()
        return
    for k in ("dlg_pengirim", "dlg_pic", "dlg_direktorat_manual", "dlg_direktorat_pilihan", "dlg_tanggal"):
        st.session_state.pop(k, None)
    dialog_buat_penerimaan()


@st.dialog("Hapus BAPP")
def dialog_konfirmasi_hapus_bapp(idx):
    item = st.session_state.scan_list[idx]
    st.write("Apakah Anda yakin ingin menghapus BAPP ini dari penerimaan?")
    st.caption(f"{item['nomor_transaksi']} — {item['nama_sekolah']}")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c2:
        if st.button("Hapus", type="primary", use_container_width=True):
            try:
                with st.spinner("Menghapus..."):
                    hapus_satu_baris_bapp(item)
                del st.session_state.scan_list[idx]
                st.session_state.scan_message = None
                st.rerun()
            except Exception as e:
                st.error(f"Gagal menghapus: {_pesan_error_ramah(e)}")


@st.dialog("Hapus Penerimaan")
def dialog_konfirmasi_hapus_penerimaan(nomor):
    st.write("Apakah Anda yakin ingin menghapus penerimaan ini?")
    st.caption(f"Nomor Penerimaan: {nomor}")
    st.warning("BAPP dalam penerimaan ini akan tersedia lagi untuk diterima ulang di penerimaan lain.")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c2:
        if st.button("Hapus", type="primary", use_container_width=True):
            try:
                with st.spinner("Menghapus penerimaan..."):
                    hapus_penerimaan(nomor)
                st.session_state.halaman = "daftar"
                st.rerun()
            except Exception as e:
                st.error(f"Gagal menghapus: {_pesan_error_ramah(e)}")


@st.dialog("Batalkan Penerimaan")
def dialog_konfirmasi_batalkan_penerimaan():
    """Membatalkan penerimaan yang masih OPEN. Karena tiap scan sudah langsung
    tertulis ke Spreadsheet, ini juga mengosongkan kembali data yang sudah
    sempat tersimpan (BAPP-nya jadi tersedia lagi untuk diterima ulang)."""
    info = st.session_state.get("penerimaan_aktif") or {}
    jumlah = len(st.session_state.scan_list)
    st.write("Apakah Anda yakin ingin membatalkan penerimaan ini?")
    if jumlah:
        st.warning(f"{jumlah} BAPP yang sudah tersimpan pada penerimaan ini akan dikosongkan kembali (bisa diterima ulang di penerimaan lain).")
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c2:
        if st.button("Ya, Batalkan", type="primary", use_container_width=True):
            try:
                if info.get("nomor_penerimaan"):
                    with st.spinner("Membatalkan..."):
                        hapus_penerimaan(info["nomor_penerimaan"])
                st.session_state.scan_list = []
                st.session_state.scan_message = None
                st.session_state.penerimaan_aktif = None
                st.session_state.halaman = "daftar"
                st.rerun()
            except Exception as e:
                st.error(f"Gagal membatalkan: {_pesan_error_ramah(e)}")


@st.dialog("Tutup Penerimaan")
def dialog_konfirmasi_tutup_penerimaan(nomor):
    st.write(
        f"Apakah Anda yakin ingin menutup penerimaan **{nomor}**? Setelah ditutup, "
        f"data tidak dapat diedit kembali dan seluruh BAPP akan ditandai sebagai DITERIMA."
    )
    c1, c2 = st.columns(2)
    with c1:
        if st.button("Batal", use_container_width=True):
            st.rerun()
    with c2:
        if st.button("Ya, Tutup", type="primary", use_container_width=True):
            try:
                with st.spinner("Menutup penerimaan..."):
                    tutup_penerimaan(nomor)
                st.session_state.scan_list = []
                st.session_state.scan_message = None
                st.session_state.penerimaan_aktif = None
                st.session_state.halaman = "detail"
                st.session_state.detail_nomor = nomor
                st.toast(f"Penerimaan {nomor} berhasil ditutup.", icon="✅")
                st.rerun()
            except Exception as e:
                st.error(f"Gagal menutup penerimaan: {_pesan_error_ramah(e)}")


# =====================================================================
# 9. INISIALISASI
# =====================================================================

init_db()

if "master_df" not in st.session_state:
    refresh_master_data()
else:
    umur_detik = (datetime.now() - st.session_state.get("last_refresh_ts", datetime.min)).total_seconds()
    if umur_detik > AUTO_REFRESH_DETIK:
        refresh_master_data()

for key, default in [
    ("scan_list", []),
    ("scan_message", None),
    ("halaman", "daftar"),
    ("daftar_halaman_ke", 1),
    ("penerimaan_aktif", None),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# =====================================================================
# 10. SIDEBAR / NAVIGASI
# =====================================================================

with st.sidebar:
    st.markdown("## 📦 Sistem BAPP")
    st.markdown("")

    st.markdown("**Penerimaan BAPP**")
    if st.button("📋 Daftar Penerimaan BAPP", use_container_width=True,
                 type="primary" if st.session_state.halaman in ("daftar", "detail") else "secondary"):
        st.session_state.halaman = "daftar"
        st.rerun()
    if st.button("➕ Buat Penerimaan Baru", use_container_width=True,
                 type="primary" if st.session_state.halaman == "form_baru" else "secondary"):
        buka_dialog_penerimaan_baru()

    st.markdown("---")
    if st.button("⚙️ Pengaturan", use_container_width=True,
                 type="primary" if st.session_state.halaman == "pengaturan" else "secondary"):
        st.session_state.halaman = "pengaturan"
        st.rerun()

    st.caption(f"Data master: {st.session_state.get('last_refresh', 'belum dimuat')}")

if st.session_state.get("load_error"):
    st.warning(f"⚠️ Gagal memuat data master dari Google Spreadsheet: {st.session_state.load_error}")


# =====================================================================
# 11. HALAMAN: DAFTAR PENERIMAAN BAPP
# =====================================================================

if st.session_state.halaman == "daftar":
    c_judul, c_tombol = st.columns([5, 2])
    with c_judul:
        st.title("Daftar Penerimaan BAPP")
        st.caption("Kelola dan pantau seluruh penerimaan BAPP")
    with c_tombol:
        st.write("")
        if st.button("+ Buat Penerimaan Baru", type="primary", use_container_width=True):
            buka_dialog_penerimaan_baru()

    df_riwayat = get_riwayat()

    f1, f2, f3, f4 = st.columns([2, 1, 1, 1])
    with f1:
        cari = st.text_input("Cari", label_visibility="collapsed", placeholder="🔎 Cari pengirim / Nomor Penerimaan")
    with f2:
        opsi_dir = ["Semua Direktorat"]
        if not df_riwayat.empty:
            opsi_dir += sorted(df_riwayat["Direktorat"].dropna().unique().tolist())
        filter_direktorat = st.selectbox("Direktorat", opsi_dir, label_visibility="collapsed")
    with f3:
        filter_tanggal = st.date_input("Tanggal", value=None, label_visibility="collapsed")
    with f4:
        filter_status = st.selectbox("Status", ["Semua Status", "🟡 OPEN", "🟢 DITERIMA"], label_visibility="collapsed")

    hasil = df_riwayat.copy()
    if cari:
        hasil = hasil[hasil.apply(lambda r: cari.lower() in str(r.values).lower(), axis=1)]
    if filter_direktorat != "Semua Direktorat":
        hasil = hasil[hasil["Direktorat"] == filter_direktorat]
    if filter_tanggal:
        hasil = hasil[hasil["Tanggal"] == filter_tanggal.strftime("%d/%m/%Y")]
    if filter_status == "🟡 OPEN":
        hasil = hasil[hasil["Status"] == STATUS_OPEN]
    elif filter_status == "🟢 DITERIMA":
        hasil = hasil[hasil["Status"] == STATUS_DITERIMA]

    st.markdown("---")

    if hasil.empty:
        st.info("Belum ada penerimaan yang cocok.")
    else:
        total_hal = max(1, -(-len(hasil) // UKURAN_HALAMAN_DAFTAR))
        halaman_ke = min(st.session_state.daftar_halaman_ke, total_hal)
        awal = (halaman_ke - 1) * UKURAN_HALAMAN_DAFTAR
        potongan = hasil.iloc[awal:awal + UKURAN_HALAMAN_DAFTAR]

        lebar = [0.4, 1.1, 0.9, 1.1, 0.8, 0.6, 0.8, 1.1, 0.4, 0.4, 0.4]
        judul_kolom = ["No", "Nomor Penerimaan", "Tanggal", "Pengirim", "Direktorat",
                        "Jumlah", "PIC", "Status", "", "", ""]
        for kolom, teks in zip(st.columns(lebar), judul_kolom):
            kolom.markdown(f"**{teks}**")

        for i, (_, baris) in enumerate(potongan.iterrows(), start=awal + 1):
            c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11 = st.columns(lebar)
            c1.write(i)
            c2.write(baris["Nomor Penerimaan"])
            c3.write(baris.get("Tanggal", "-"))
            c4.write(baris.get("Pengirim") or "-")
            c5.markdown(render_badge(baris.get("Direktorat", "-"), "biru"), unsafe_allow_html=True)
            c6.write(int(baris["Jumlah BAPP"]))
            c7.write(baris.get("PIC") or "-")

            status = baris.get("Status", STATUS_DITERIMA)
            sedang_open = status == STATUS_OPEN
            if sedang_open:
                c8.markdown(render_badge("🟡 OPEN", "kuning"), unsafe_allow_html=True)
                if c9.button("✏️", key=f"edit_{baris['Nomor Penerimaan']}", help="Lanjutkan penerimaan"):
                    berhasil, pesan_error = muat_penerimaan_open(baris["Nomor Penerimaan"])
                    if berhasil:
                        st.session_state.halaman = "form_baru"
                        st.rerun()
                    else:
                        st.error(pesan_error)
            else:
                c8.markdown(render_badge("🟢 DITERIMA", "hijau"), unsafe_allow_html=True)

            if c10.button("👁", key=f"detail_{baris['Nomor Penerimaan']}", help="Lihat detail"):
                st.session_state.halaman = "detail"
                st.session_state.detail_nomor = baris["Nomor Penerimaan"]
                st.rerun()

            if sedang_open:
                if c11.button("🗑", key=f"hapusdaftar_{baris['Nomor Penerimaan']}", help="Hapus penerimaan"):
                    dialog_konfirmasi_hapus_penerimaan(baris["Nomor Penerimaan"])

        st.markdown("---")
        cp1, cp2, cp3 = st.columns([1, 2, 1])
        with cp1:
            if st.button("← Sebelumnya", disabled=halaman_ke <= 1):
                st.session_state.daftar_halaman_ke = halaman_ke - 1
                st.rerun()
        with cp2:
            st.markdown(
                f"<div style='text-align:center;color:#6b7280;'>Halaman {halaman_ke} dari {total_hal}</div>",
                unsafe_allow_html=True,
            )
        with cp3:
            if st.button("Berikutnya →", disabled=halaman_ke >= total_hal):
                st.session_state.daftar_halaman_ke = halaman_ke + 1
                st.rerun()


# =====================================================================
# 12. HALAMAN: SCAN BAPP (setelah popup Buat Penerimaan Baru)
# =====================================================================

elif st.session_state.halaman == "form_baru":
    info = st.session_state.get("penerimaan_aktif")

    if not info:
        st.warning("Belum ada penerimaan aktif. Klik tombol di bawah untuk memulai.")
        if st.button("+ Buat Penerimaan Baru", type="primary"):
            buka_dialog_penerimaan_baru()
    else:
        if st.button("← Kembali ke Daftar Penerimaan"):
            st.session_state.halaman = "daftar"
            st.rerun()

        c_judul, c_status = st.columns([4, 1])
        with c_judul:
            st.markdown(f"### PENERIMAAN BAPP — {info['nomor_penerimaan']}")
        with c_status:
            st.markdown(render_badge("🟡 OPEN / Belum Diterima", "kuning"), unsafe_allow_html=True)

        c1, c2, c3, c4, c5 = st.columns(5)
        c1.markdown(f"**Pengirim**  \n{info['nama_pengirim']}")
        c2.markdown(f"**Direktorat**  \n{info['direktorat']}")
        c3.markdown(f"**Tanggal**  \n{format_tanggal_indonesia(info['tanggal'])}")
        c4.markdown(f"**PIC**  \n{info['pic']}")
        c5.markdown(f"**Jumlah BAPP**  \n{len(st.session_state.scan_list)}")

        st.markdown("---")
        st.text_input(
            "🔍 Scan Barcode / Ketik Nomor Transaksi",
            key="input_scan",
            placeholder="Scan barcode atau ketik Nomor Transaksi lalu tekan Enter...",
            on_change=handle_scan_input,
        )
        autofocus_scan_input()

        if st.session_state.scan_message:
            tipe, pesan = st.session_state.scan_message
            getattr(st, tipe)(pesan)

        st.markdown("#### Daftar BAPP")
        st.caption("Setiap BAPP yang berhasil di-scan langsung tersimpan (status OPEN) -- aman dilanjutkan nanti.")
        if st.session_state.scan_list:
            lebar = [0.5, 1.4, 1, 1.7, 1.4, 0.9, 1.1, 1.4, 0.5]
            judul = ["No", "Nomor Transaksi", "NPSN", "Nama Sekolah", "Nomor Penerimaan Pertama",
                     "Nomor Urut", "Serial Number", "Nama Koordinator", ""]
            for kolom, teks in zip(st.columns(lebar), judul):
                kolom.markdown(f"**{teks}**")

            for i, item in enumerate(st.session_state.scan_list):
                c1, c2, c3, c4, c5, c6, c7, c8, c9 = st.columns(lebar)
                c1.write(i + 1)
                c2.write(item["nomor_transaksi"])
                c3.write(item["npsn"])
                c4.write(item["nama_sekolah"])
                c5.write(ekstrak_nomor_penerimaan_pertama(item.get("nomor_penerimaan_pertama", "")) or "-")
                c6.write(item.get("nomor_urut_pertama", "") or "-")
                c7.write(item["serial_number"])
                c8.write(item["nama_koordinator"])
                if c9.button("🗑", key=f"hapus_scan_{i}"):
                    dialog_konfirmasi_hapus_bapp(i)
        else:
            st.info("Belum ada BAPP yang di-scan.")

        st.markdown("---")
        col_tutup, col_batal = st.columns([2, 1])
        with col_tutup:
            if st.button("🔒 CLOSE FOLDER / SIMPAN PENERIMAAN", type="primary", use_container_width=True,
                         disabled=len(st.session_state.scan_list) == 0):
                dialog_konfirmasi_tutup_penerimaan(info["nomor_penerimaan"])
        with col_batal:
            if st.button("❌ Batalkan Penerimaan", use_container_width=True):
                dialog_konfirmasi_batalkan_penerimaan()
        if not st.session_state.scan_list:
            st.caption("Scan minimal 1 BAPP sebelum bisa menutup penerimaan.")


# =====================================================================
# 13. HALAMAN: DETAIL PENERIMAAN
# =====================================================================

elif st.session_state.halaman == "detail":
    if st.button("← Kembali"):
        st.session_state.halaman = "daftar"
        st.rerun()

    nomor = st.session_state.get("detail_nomor")
    tabel, info = get_detail_penerimaan(nomor) if nomor else (pd.DataFrame(), {})

    if not info:
        st.warning("Data penerimaan tidak ditemukan. Coba refresh data master di menu Pengaturan.")
    else:
        c_judul, c_status = st.columns([4, 1])
        with c_judul:
            st.title("DETAIL PENERIMAAN")
        with c_status:
            st.write("")
            if info["status"] == STATUS_OPEN:
                st.markdown(render_badge("🟡 OPEN / Belum Diterima", "kuning"), unsafe_allow_html=True)
            else:
                st.markdown(render_badge("🟢 DITERIMA", "hijau"), unsafe_allow_html=True)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**Nomor Penerimaan**")
            st.write(info["nomor_penerimaan"])
            st.markdown("**Tanggal**")
            st.write(info["waktu"].split(" ")[0] if info["waktu"] else "-")
        with c2:
            st.markdown("**Pengirim**")
            st.write(info["pengirim"] or "-")
            st.markdown("**Direktorat**")
            st.markdown(render_badge(info["direktorat"] or "-", "biru"), unsafe_allow_html=True)
        with c3:
            st.markdown("**PIC Penerimaan**")
            st.write(info["pic"] or "-")
            st.markdown("**Jumlah BAPP**")
            st.write(info["jumlah"])

        st.markdown("---")
        st.markdown("#### Daftar BAPP")
        st.dataframe(tabel, use_container_width=True, hide_index=True)

        st.markdown("---")
        if info["status"] == STATUS_OPEN:
            st.info("Penerimaan ini masih **OPEN** -- print baru tersedia setelah ditutup (semua BAPP ditandai DITERIMA).")
            if st.button("✏️ Lanjutkan Penerimaan Ini", type="primary"):
                berhasil, pesan_error = muat_penerimaan_open(nomor)
                if berhasil:
                    st.session_state.halaman = "form_baru"
                    st.rerun()
                else:
                    st.error(pesan_error)
        else:
            with st.spinner("Menyiapkan PDF..."):
                pdf_bytes = buat_pdf_penerimaan(info, tabel)
            render_tombol_pdf(pdf_bytes, f"BAPP_{info['nomor_penerimaan']}.pdf")


# =====================================================================
# 14. HALAMAN: PENGATURAN
# =====================================================================

elif st.session_state.halaman == "pengaturan":
    st.title("⚙️ Pengaturan")

    st.subheader("Data Master")
    if st.button("🔄 Refresh Data dari Spreadsheet"):
        with st.spinner("Memuat ulang data..."):
            ok = refresh_master_data()
        if ok:
            st.success("Data master berhasil diperbarui.")
        else:
            st.error(f"Gagal mengambil data: {st.session_state.load_error}")

    st.markdown("---")
    st.subheader("Format Nomor Penerimaan")
    termin_sekarang = get_setting("termin_penerimaan", "2")
    termin_baru = st.text_input(
        "Angka termin penerimaan (contoh: '2' akan menghasilkan SD2-001)", value=termin_sekarang,
    )
    if st.button("Simpan Pengaturan"):
        set_setting("termin_penerimaan", termin_baru.strip())
        st.success("Pengaturan disimpan. Nomor penerimaan berikutnya akan memakai termin ini.")

    st.markdown("---")
    st.subheader("Koneksi Spreadsheet")
    st.write(f"Spreadsheet ID: `{SPREADSHEET_ID}`")
    st.write(f"Nama Sheet: `{SHEET_NAME}`")
    st.caption(
        "Service account harus memiliki akses **Editor** (bukan Viewer) karena "
        "aplikasi ini menulis balik ke sheet 'data'."
    )
    st.caption("Kolom yang wajib ada di sheet 'data': " + ", ".join(KOLOM_WAJIB))
    if st.button("🔌 Test Koneksi ke Spreadsheet"):
        with st.spinner("Menguji koneksi..."):
            ok = refresh_master_data()
        if ok:
            st.success("Koneksi berhasil! Data master terbaca dengan baik.")
        else:
            st.error(f"Koneksi gagal: {st.session_state.load_error}")
