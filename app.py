"""
SISTEM PENERIMAAN BAPP
=======================
Aplikasi Streamlit untuk mencatat penerimaan BAPP fisik menggunakan
scanner barcode. Data master DAN hasil penerimaan sama-sama disimpan
di Google Spreadsheet (sheet "data").

Flow: Dashboard -> Daftar Penerimaan BAPP -> (+) Buat Penerimaan Baru
(popup info) -> halaman scan BAPP -> Simpan -> Detail -> Print.

Cara menjalankan:
    streamlit run app.py

Lihat PANDUAN.md untuk instruksi instalasi & konfigurasi lengkap.
"""

import os
import io
import json
import time
from datetime import datetime

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import sqlite3

import gspread
from gspread.utils import rowcol_to_a1
from google.oauth2.service_account import Credentials

from reportlab.lib.units import cm
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
BAPP_PER_LEMBAR_PRINT = 40

# Ukuran kertas custom untuk print Penerimaan BAPP: 24 cm x 28 cm, portrait.
KERTAS_PRINT = (24 * cm, 28 * cm)

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
    ringkasan = ringkasan.sort_values("_waktu_dt", ascending=False).drop(columns=["_waktu_dt", "Waktu"])

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

    buffer = io.BytesIO()

    # ---------------------------------------------------------------
    # UKURAN KERTAS DAN MARGIN
    # ---------------------------------------------------------------

    margin = 1 * cm

    lebar_isi = KERTAS_PRINT[0] - 2 * margin

    doc = SimpleDocTemplate(
        buffer,
        pagesize=KERTAS_PRINT,

        topMargin=1 * cm,
        bottomMargin=1.3 * cm,
        leftMargin=margin,
        rightMargin=margin,
    )

    styles = getSampleStyleSheet()


    # ---------------------------------------------------------------
    # STYLE JUDUL
    # ---------------------------------------------------------------

    judul_style = ParagraphStyle(
        "Judul",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=14,
        spaceAfter=4,
        alignment=1,
    )


    # ---------------------------------------------------------------
    # STYLE ISI TABEL
    # ---------------------------------------------------------------
    # alignment=1  -> tengah horizontal
    # leading=6     -> jarak antarbaris
    # spaceBefore=0
    # spaceAfter=0

    sel_style = ParagraphStyle(
        "Sel",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=6,
        leading=6,
        alignment=1,
        spaceBefore=0,
        spaceAfter=0,
    )


    # ---------------------------------------------------------------
    # STYLE HEADER TABEL
    # ---------------------------------------------------------------

    header_sel_style = ParagraphStyle(
        "HeaderSel",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7,
        leading=7,
        alignment=1,
        spaceBefore=0,
        spaceAfter=0,
        textColor=colors.HexColor("#111827"),
    )


    # ---------------------------------------------------------------
    # FUNGSI UNTUK SEMUA ISI CELL
    # ---------------------------------------------------------------

    def sel(txt):

        nilai = str(txt).strip()

        if not nilai:
            nilai = "-"

        return Paragraph(
            nilai,
            sel_style
        )


    # ---------------------------------------------------------------
    # FUNGSI UNTUK HEADER
    # ---------------------------------------------------------------

    def header_sel(txt):

        return Paragraph(
            str(txt),
            header_sel_style
        )


    # ---------------------------------------------------------------
    # AMBIL TANGGAL
    # ---------------------------------------------------------------

    tanggal_saja = (
        info["waktu"].split(" ")[0]
        if info.get("waktu")
        else "-"
    )


    # =================================================================
    # LEBAR KOLOM
    # =================================================================
    #
    # TOTAL = 22 CM
    #
    # Dibuat ulang supaya:
    # - Barcode Penerimaan tidak terlalu sempit
    # - Nomor Penerimaan tidak terlalu sempit
    # - Serial Number tetap cukup
    # - Nama Koordinator cukup
    #
    # =================================================================

    lebar_kolom = [
        0.8 * cm,   # 1. No
        2.5 * cm,   # 2. Nomor Transaksi
        1.6 * cm,   # 3. NPSN
        3.5 * cm,   # 4. Nama Sekolah
        1.7 * cm,   # 5. Tanggal BAPP
        2.1 * cm,   # 6. Barcode Penerimaan
        2.0 * cm,   # 7. Nomor Penerimaan 1
        1.3 * cm,   # 8. Nomor Urut
        3.5 * cm,   # 9. Serial Number
        3.0 * cm,   # 10. Nama Koordinator
    ]


    # =================================================================
    # FLOW PDF
    # =================================================================

    flow = []

    total_baris = len(tabel_df)

    total_lembar = max(
        1,
        -(-total_baris // BAPP_PER_LEMBAR_PRINT)
    )


    # =================================================================
    # LOOP PER HALAMAN
    # =================================================================

    for lembar in range(total_lembar):

        potongan = tabel_df.iloc[
            lembar * BAPP_PER_LEMBAR_PRINT:
            (lembar + 1) * BAPP_PER_LEMBAR_PRINT
        ]


        # -------------------------------------------------------------
        # JUDUL
        # -------------------------------------------------------------

        flow.append(
            Paragraph(
                "BUKTI PENERIMAAN BAPP",
                judul_style
            )
        )


        # =============================================================
        # BLOK INFO
        # =============================================================

        info_rows = [

            [
                "Nomor Penerimaan",
                ":",
                info.get("nomor_penerimaan") or "-",

                "Tanggal",
                ":",
                tanggal_saja
            ],

            [
                "Pengirim",
                ":",
                info.get("pengirim") or "-",

                "PIC Penerimaan",
                ":",
                info.get("pic") or "-"
            ],

            [
                "Direktorat",
                ":",
                info.get("direktorat") or "-",

                "Jumlah BAPP",
                ":",
                str(info.get("jumlah", ""))
            ],
        ]


        t_info = Table(
            info_rows,
            colWidths=[
                3.2 * cm,
                0.4 * cm,
                7.6 * cm,
                3.2 * cm,
                0.4 * cm,
                7.2 * cm
            ]
        )


        t_info.setStyle(
            TableStyle([

                ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),

                ("FONTSIZE", (0, 0), (-1, -1), 9.5),

                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),

                ("FONTNAME", (3, 0), (3, -1), "Helvetica-Bold"),

                ("TOPPADDING", (0, 0), (-1, -1), 1),

                ("BOTTOMPADDING", (0, 0), (-1, -1), 1),

                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

            ])
        )


        flow.append(t_info)

        flow.append(
            Spacer(
                1,
                0.2 * cm
            )
        )


        # =============================================================
        # HEADER TABEL
        # =============================================================

        header_tabel = [

            header_sel(t)

            for t in [

                "No",

                "Nomor Transaksi",

                "NPSN",

                "Nama Sekolah",

                "Tanggal BAPP",

                "Barcode Penerimaan",

                "Nomor Penerimaan 1",

                "Nomor Urut",

                "Serial Number",

                "Nama Koordinator",
            ]
        ]


        data_tabel = [
            header_tabel
        ]


        # =============================================================
        # ISI TABEL
        # =============================================================
        #
        # PENTING:
        # SEMUA kolom sekarang menggunakan sel()
        #
        # Jadi tidak ada lagi:
        # str(r["NPSN"])
        # str(r["Nomor"])
        # str(r["Tanggal BAPP"])
        # str(r["Nomor Urut"])
        #
        # Semuanya Paragraph -> posisi vertical lebih rapi.
        # =============================================================

        for _, r in potongan.iterrows():

            data_tabel.append([

                # 1. No
                sel(r["Nomor"]),

                # 2. Nomor Transaksi
                sel(r["Nomor Transaksi"]),

                # 3. NPSN
                sel(r["NPSN"]),

                # 4. Nama Sekolah
                sel(r["Nama Sekolah"]),

                # 5. Tanggal BAPP
                sel(r["Tanggal BAPP"]),

                # 6. Barcode Penerimaan
                buat_sel_barcode(
                    r["Barcode Penerimaan"],
                    style_teks=sel_style
                ),

                # 7. Nomor Penerimaan 1
                sel(r["Nomor Penerimaan Pertama"]),

                # 8. Nomor Urut
                sel(r["Nomor Urut"]),

                # 9. Serial Number
                sel(r["Serial Number"]),

                # 10. Nama Koordinator
                sel(r["Nama Koordinator"]),
            ])


        # =============================================================
        # BUAT TABEL
        # =============================================================

        t = Table(
            data_tabel,
            colWidths=lebar_kolom,
            repeatRows=1,

            # Membuat tinggi baris isi lebih konsisten
            rowHeights=[
                None
            ] + [
                0.55 * cm
                for _ in range(len(potongan))
            ]
        )


        # =============================================================
        # STYLE TABEL
        # =============================================================

        t.setStyle(
            TableStyle([

                # -----------------------------------------------------
                # HEADER
                # -----------------------------------------------------

                (
                    "BACKGROUND",
                    (0, 0),
                    (-1, 0),
                    colors.HexColor("#e5e7eb")
                ),


                # -----------------------------------------------------
                # GARIS TABEL
                # -----------------------------------------------------

                (
                    "GRID",
                    (0, 0),
                    (-1, -1),
                    0.5,
                    colors.HexColor("#111827")
                ),


                # -----------------------------------------------------
                # FONT ISI
                # -----------------------------------------------------

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
                    6
                ),


                # -----------------------------------------------------
                # JARAK DALAM CELL
                # -----------------------------------------------------

                (
                    "TOPPADDING",
                    (0, 0),
                    (-1, -1),
                    1
                ),

                (
                    "BOTTOMPADDING",
                    (0, 0),
                    (-1, -1),
                    1
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


                # -----------------------------------------------------
                # SEMUA ISI TENGAH HORIZONTAL
                # -----------------------------------------------------

                (
                    "ALIGN",
                    (0, 0),
                    (-1, -1),
                    "CENTER"
                ),


                # -----------------------------------------------------
                # SEMUA ISI TENGAH VERTICAL
                # -----------------------------------------------------

                (
                    "VALIGN",
                    (0, 0),
                    (-1, -1),
                    "MIDDLE"
                ),

            ])
        )


        flow.append(t)


        flow.append(
            Spacer(
                1,
                0.4 * cm
            )
        )


        # =============================================================
        # TANDA TANGAN
        # =============================================================

        pengirim_label = (
            info.get("pengirim")
            or "..........................."
        )

        pic_label = (
            info.get("pic")
            or "..........................."
        )


        ttd_rows = [

            [
                "Pengirim",
                "",
                "PIC Penerimaan"
            ],

            [
                "",
                "",
                ""
            ],

            [
                "",
                "",
                ""
            ],

            [
                f"({pengirim_label})",
                "",
                f"({pic_label})"
            ],
        ]


        t_ttd = Table(

            ttd_rows,

            colWidths=[
                7 * cm,
                lebar_isi - 14 * cm,
                7 * cm
            ],

            rowHeights=[
                0.45 * cm,
                0.9 * cm,
                0.15 * cm,
                0.45 * cm
            ]
        )


        t_ttd.setStyle(
            TableStyle([

                (
                    "FONTNAME",
                    (0, 0),
                    (-1, -1),
                    "Helvetica"
                ),

                (
                    "FONTSIZE",
                    (0, 0),
                    (-1, -1),
                    10
                ),

                (
                    "FONTNAME",
                    (0, 0),
                    (0, 0),
                    "Helvetica-Bold"
                ),

                (
                    "FONTNAME",
                    (2, 0),
                    (2, 0),
                    "Helvetica-Bold"
                ),

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


        flow.append(t_ttd)


        # =============================================================
        # PAGE BREAK
        # =============================================================

        if lembar < total_lembar - 1:

            flow.append(
                PageBreak()
            )


    # =================================================================
    # BUILD PDF
    # =================================================================

    doc.build(
        flow,
        canvasmaker=NumberedCanvas
    )


    buffer.seek(0)

    return buffer.getvalue()

# =====================================================================
# 7. KOMPONEN TAMPILAN KECIL (card, badge)
# =====================================================================

def render_metric_card(label, value, icon=""):
    st.markdown(
        f"""
        <div style="background:#ffffff;border-radius:14px;padding:18px 20px;
                    box-shadow:0 1px 3px rgba(0,0,0,0.08);border:1px solid #eef0f2;">
            <div style="font-size:13px;color:#6b7280;margin-bottom:6px;">{icon} {label}</div>
            <div style="font-size:26px;font-weight:700;color:#111827;">{value}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


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
    ("halaman", "dashboard"),
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

    if st.button("📊 Dashboard", use_container_width=True,
                 type="primary" if st.session_state.halaman == "dashboard" else "secondary"):
        st.session_state.halaman = "dashboard"
        st.rerun()

    st.markdown("**Penerimaan BAPP**")
    if st.button("📋 Daftar Penerimaan BAPP", use_container_width=True,
                 type="primary" if st.session_state.halaman in ("daftar", "detail") else "secondary"):
        st.session_state.halaman = "daftar"
        st.rerun()
    if st.button("➕ Buat Penerimaan Baru", use_container_width=True,
                 type="primary" if st.session_state.halaman == "form_baru" else "secondary"):
        buka_dialog_penerimaan_baru()

    st.markdown("")
    if st.button("📈 Laporan", use_container_width=True,
                 type="primary" if st.session_state.halaman == "laporan" else "secondary"):
        st.session_state.halaman = "laporan"
        st.rerun()

    st.markdown("---")
    if st.button("⚙️ Pengaturan", use_container_width=True,
                 type="primary" if st.session_state.halaman == "pengaturan" else "secondary"):
        st.session_state.halaman = "pengaturan"
        st.rerun()

    st.caption(f"Data master: {st.session_state.get('last_refresh', 'belum dimuat')}")

if st.session_state.get("load_error"):
    st.warning(f"⚠️ Gagal memuat data master dari Google Spreadsheet: {st.session_state.load_error}")


# =====================================================================
# 11. HALAMAN: DASHBOARD
# =====================================================================

if st.session_state.halaman == "dashboard":
    c_judul, c_refresh = st.columns([6, 1])
    with c_judul:
        st.title("Dashboard Penerimaan BAPP")
        st.caption("Pantau progres penerimaan BAPP secara realtime")
    with c_refresh:
        st.write("")
        if st.button("🔄 Refresh"):
            with st.spinner("Memuat ulang data..."):
                ok = refresh_master_data()
            if ok:
                st.success("Data diperbarui.")
            else:
                st.error(st.session_state.load_error)

    df_riwayat = get_riwayat()
    df_master = st.session_state.get("master_df", pd.DataFrame())

    total_penerimaan = len(df_riwayat)
    total_bapp_keseluruhan = len(df_master)

    df_diterima_master = pd.DataFrame()
    if not df_master.empty and KOLOM_STATUS_BARU in df_master.columns:
        df_diterima_master = df_master[
            df_master[KOLOM_STATUS_BARU].astype(str).str.strip().str.upper() == STATUS_DITERIMA
        ]
    total_bapp = len(df_diterima_master)
    persen_diterima_keseluruhan = (total_bapp / total_bapp_keseluruhan * 100) if total_bapp_keseluruhan else 0

    hari_ini_str = datetime.now().strftime("%d/%m/%Y")
    bapp_hari_ini_df = pd.DataFrame()
    if not df_master.empty and KOLOM_WAKTU_BARU in df_master.columns:
        bapp_hari_ini_df = df_master[df_master[KOLOM_WAKTU_BARU].astype(str).str.startswith(hari_ini_str)]
    total_hari_ini = len(bapp_hari_ini_df)

    c1, c2, c3 = st.columns(3)
    with c1:
        render_metric_card("Total Penerimaan", total_penerimaan, "📥")
    with c2:
        render_metric_card(
            "Total BAPP Diterima",
            f"{total_bapp:,}".replace(",", ".") + f" ({persen_diterima_keseluruhan:.1f}%)",
            "📦",
        )
    with c3:
        render_metric_card("Penerimaan Hari Ini", f"{total_hari_ini} BAPP", "📅")
    st.caption(f"Persentase dihitung dari total {total_bapp_keseluruhan:,} BAPP di data master.".replace(",", "."))

    st.markdown("")
    st.markdown("#### Penerimaan per Direktorat")
    st.caption("Belum Discan = belum pernah di-scan sama sekali di aplikasi ini.")
    if df_master.empty or "Direktorat" not in df_master.columns:
        st.info("Belum ada data untuk breakdown per Direktorat.")
    else:
        def _status_bucket(v):
            v = str(v).strip().upper()
            if v == STATUS_OPEN:
                return "OPEN"
            if v == STATUS_DITERIMA:
                return "DITERIMA"
            return "Belum Discan"

        df_pivot_src = df_master.copy()
        if KOLOM_STATUS_BARU in df_pivot_src.columns:
            df_pivot_src["_status_bucket"] = df_pivot_src[KOLOM_STATUS_BARU].apply(_status_bucket)
        else:
            df_pivot_src["_status_bucket"] = "Belum Discan"

        pivot = df_pivot_src.groupby(["Direktorat", "_status_bucket"]).size().unstack(fill_value=0)
        for kolom in ["Belum Discan", "OPEN", "DITERIMA"]:
            if kolom not in pivot.columns:
                pivot[kolom] = 0
        pivot = pivot[["Belum Discan", "OPEN", "DITERIMA"]]
        pivot["Grand Total"] = pivot.sum(axis=1)
        pivot = pivot.sort_index()

        baris_total = pivot.sum(axis=0)
        baris_total.name = "Grand Total"
        pivot = pd.concat([pivot, baris_total.to_frame().T])

        pivot = pivot.reset_index().rename(columns={"index": "Direktorat"})
        for kolom in ["Belum Discan", "OPEN", "DITERIMA", "Grand Total"]:
            pivot[kolom] = pivot[kolom].astype(int)

        st.dataframe(pivot, use_container_width=True, hide_index=True)

    st.markdown("")
    col_terbaru, col_ringkasan = st.columns([1.4, 1])

    with col_terbaru:
        st.markdown("#### Penerimaan Terbaru")
        if df_riwayat.empty:
            st.info("Belum ada penerimaan.")
        else:
            lebar_terbaru = [1.2, 1.1, 0.8, 0.8, 1, 0.5]
            judul_terbaru = ["Nomor Penerimaan", "Pengirim", "Direktorat", "Jumlah BAPP", "Status", ""]
            for kolom, teks in zip(st.columns(lebar_terbaru), judul_terbaru):
                kolom.markdown(f"**{teks}**")
            for _, baris in df_riwayat.head(5).iterrows():
                c1, c2, c3, c4, c5, c6 = st.columns(lebar_terbaru)
                c1.write(baris["Nomor Penerimaan"])
                c2.write(baris.get("Pengirim") or "-")
                c3.markdown(render_badge(baris["Direktorat"], "biru"), unsafe_allow_html=True)
                c4.write(int(baris["Jumlah BAPP"]))
                if baris.get("Status") == STATUS_OPEN:
                    c5.markdown(render_badge("🟡 OPEN", "kuning"), unsafe_allow_html=True)
                else:
                    c5.markdown(render_badge("🟢 DITERIMA", "hijau"), unsafe_allow_html=True)
                if c6.button("👁", key=f"terbaru_{baris['Nomor Penerimaan']}"):
                    st.session_state.halaman = "detail"
                    st.session_state.detail_nomor = baris["Nomor Penerimaan"]
                    st.rerun()
            if st.button("Lihat Semua →"):
                st.session_state.halaman = "daftar"
                st.rerun()

    with col_ringkasan:
        st.markdown("#### Ringkasan Hari Ini")
        render_metric_card("Total Hari Ini", f"{total_hari_ini} BAPP", "📦")
        st.markdown("")
        if not bapp_hari_ini_df.empty and "Direktorat" in bapp_hari_ini_df.columns:
            rekap_hari_ini = (
                bapp_hari_ini_df.groupby("Direktorat").size()
                .reset_index(name="Jumlah").sort_values("Jumlah", ascending=False)
            )
            for _, baris in rekap_hari_ini.iterrows():
                st.markdown(f"{baris['Direktorat']} — **{baris['Jumlah']} BAPP**")
        else:
            st.caption("Belum ada aktivitas hari ini.")


# =====================================================================
# 12. HALAMAN: DAFTAR PENERIMAAN BAPP
# =====================================================================

elif st.session_state.halaman == "daftar":
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
# 13. HALAMAN: SCAN BAPP (setelah popup Buat Penerimaan Baru)
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
# 14. HALAMAN: DETAIL PENERIMAAN
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
            st.download_button(
                "🖨 Print Penerimaan (Unduh PDF)",
                data=pdf_bytes,
                file_name=f"BAPP_{info['nomor_penerimaan']}.pdf",
                mime="application/pdf",
                type="primary",
            )


# =====================================================================
# 15. HALAMAN: LAPORAN
# =====================================================================

elif st.session_state.halaman == "laporan":
    st.title("Laporan")
    st.caption("Rekap seluruh penerimaan BAPP")

    df_riwayat = get_riwayat()
    if df_riwayat.empty:
        st.info("Belum ada data penerimaan untuk dilaporkan.")
    else:
        st.markdown("#### Ringkasan per Direktorat")
        ringkas = (
            df_riwayat.groupby("Direktorat")["Jumlah BAPP"].sum()
            .reset_index().sort_values("Jumlah BAPP", ascending=False)
        )
        st.dataframe(ringkas, use_container_width=True, hide_index=True)

        st.markdown("#### Semua Penerimaan")
        st.dataframe(df_riwayat, use_container_width=True, hide_index=True)

        csv_bytes = df_riwayat.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ Unduh Laporan (CSV)", data=csv_bytes,
            file_name="laporan_penerimaan_bapp.csv", mime="text/csv",
        )


# =====================================================================
# 16. HALAMAN: PENGATURAN
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
