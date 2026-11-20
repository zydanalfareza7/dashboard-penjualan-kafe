"""
Dashboard Analitik Visual Interaktif - Penjualan Kafe
Kelompok 04 | Analisis Data

Fitur:
  * Sumber data otomatis dari Google Sheets (atau fallback ke Excel lokal)
  * Input data penjualan langsung dari web & otomatis tersimpan ke Google Sheets
  * Agregasi data tabular historis (per batch, minggu, produk)
  * Visualisasi grafik tren penjualan yang interaktif (Plotly)
  * Prediksi penjualan dengan Deep Learning (TensorFlow / Keras)

Jalankan:  streamlit run app.py
"""
import re
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from gsheets_helper import (
    APPS_SCRIPT_TEMPLATE,
    build_upserted_sales_df,
    connect_spreadsheet,
    get_gsheet_config,
    is_apps_script_url,
    read_apps_script_tables,
    read_gsheet_tables,
    save_local_config,
    write_apps_script_tables,
    write_full_tables,
)

# --------------------------------------------------------------------------
# Konfigurasi & Konstanta
# --------------------------------------------------------------------------
BASE_DIR = Path(__file__).parent
FILE_DATA = BASE_DIR / "Data_Penjualan_Kopi_1.xlsx"

METRIK = {
    "Qty": "Jumlah Terjual (cup)",
    "Omzet": "Omzet (Rp)",
    "Laba": "Laba Kotor (Rp)",
}
COFFEE_PALETTE = [
    "#6F4E37",
    "#C68B59",
    "#2E6F6D",
    "#D97706",
    "#8C3B2B",
    "#4A6B82",
    "#7C6A58",
]

st.set_page_config(
    page_title="Dashboard Analitik Penjualan Kafe",
    page_icon="☕",
    layout="wide",
    initial_sidebar_state="expanded",
)


def inject_custom_css():
    st.markdown(
        """
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700;800&display=swap');

        html, body, [class*="css"] {
            font-family: 'Plus Jakarta Sans', -apple-system, BlinkMacSystemFont, sans-serif;
        }

        /* Sembunyikan tombol Fork/GitHub/Menu kanan atas agar tampilan lebih bersih */
        [data-testid="stToolbarActions"], #MainMenu {
            visibility: hidden;
            height: 0;
        }

        .block-container {
            padding-top: 1.6rem !important;
            padding-bottom: 2.5rem !important;
            max-width: 1280px;
        }

        /* Keyframes animasi halus */
        @keyframes fadeSlideUp {
            from { opacity: 0; transform: translateY(14px); }
            to { opacity: 1; transform: translateY(0); }
        }
        @keyframes gradientShift {
            0% { background-position: 0% 50%; }
            50% { background-position: 100% 50%; }
            100% { background-position: 0% 50%; }
        }
        @keyframes pulseGlow {
            0%, 100% { box-shadow: 0 0 0 0 rgba(74, 222, 128, 0.35); }
            50% { box-shadow: 0 0 0 8px rgba(74, 222, 128, 0); }
        }

        /* White card containers in main area */
        section.main [data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border: 1px solid #EAE0D5 !important;
            border-radius: 16px !important;
            padding: 1.15rem 1.25rem !important;
            box-shadow: 0 4px 18px rgba(44, 30, 22, 0.04);
            margin-bottom: 0.75rem;
            animation: fadeSlideUp 0.45s ease-out both;
            transition: box-shadow 0.2s ease, transform 0.2s ease;
        }
        section.main [data-testid="stVerticalBlockBorderWrapper"]:hover {
            box-shadow: 0 8px 24px rgba(44, 30, 22, 0.07);
        }

        /* Hero header */
        .hero-card {
            background: linear-gradient(135deg, #2D1B14 0%, #5C3D2E 48%, #9C6B43 85%, #4A2C20 100%);
            background-size: 200% 200%;
            animation: gradientShift 10s ease infinite, fadeSlideUp 0.4s ease-out both;
            padding: 1.75rem 2.1rem;
            border-radius: 18px;
            color: #FFFDF9;
            margin-bottom: 1.35rem;
            box-shadow: 0 12px 30px rgba(45, 27, 20, 0.16);
            display: flex;
            justify-content: space-between;
            align-items: center;
            flex-wrap: wrap;
            gap: 1rem;
            border: 1px solid rgba(255, 255, 255, 0.1);
        }
        .hero-badge {
            display: inline-block;
            background: rgba(255, 255, 255, 0.14);
            border: 1px solid rgba(255, 255, 255, 0.24);
            padding: 0.25rem 0.8rem;
            border-radius: 999px;
            font-size: 0.74rem;
            font-weight: 700;
            letter-spacing: 0.06em;
            margin-bottom: 0.5rem;
        }
        .hero-title {
            font-size: 1.85rem;
            font-weight: 800;
            margin: 0;
            line-height: 1.2;
            letter-spacing: -0.02em;
        }
        .hero-sub {
            font-size: 0.92rem;
            opacity: 0.86;
            margin-top: 0.4rem;
        }
        .status-pill {
            padding: 0.5rem 1.05rem;
            border-radius: 999px;
            font-size: 0.82rem;
            font-weight: 600;
            backdrop-filter: blur(8px);
        }
        .status-gsheet {
            background: rgba(34, 197, 94, 0.22);
            border: 1px solid rgba(134, 239, 172, 0.5);
            color: #ECFDF5;
            animation: pulseGlow 2.4s ease-in-out infinite;
        }
        .status-excel {
            background: rgba(251, 191, 36, 0.2);
            border: 1px solid rgba(253, 224, 71, 0.45);
            color: #FEFCE8;
        }

        /* Custom KPI cards */
        .kpi-grid {
            display: grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap: 1rem;
            margin: 0.75rem 0 1.25rem 0;
        }
        @media (max-width: 900px) {
            .kpi-grid { grid-template-columns: repeat(2, minmax(0, 1fr)); }
        }
        .kpi-card {
            background: #FFFFFF;
            border: 1px solid #EAE0D5;
            border-radius: 16px;
            padding: 1.15rem 1.25rem;
            box-shadow: 0 4px 16px rgba(44, 30, 22, 0.045);
            position: relative;
            overflow: hidden;
            animation: fadeSlideUp 0.45s ease-out both;
            transition: transform 0.18s ease, box-shadow 0.18s ease;
        }
        .kpi-card:nth-child(1) { animation-delay: 0.04s; }
        .kpi-card:nth-child(2) { animation-delay: 0.10s; }
        .kpi-card:nth-child(3) { animation-delay: 0.16s; }
        .kpi-card:nth-child(4) { animation-delay: 0.22s; }
        .kpi-card:hover {
            transform: translateY(-3px);
            box-shadow: 0 10px 24px rgba(44, 30, 22, 0.09);
        }

        /* Hover & press animasi pada tombol dan tab */
        .stButton > button, .stFormSubmitButton > button {
            transition: transform 0.15s ease, box-shadow 0.15s ease !important;
        }
        .stButton > button:hover, .stFormSubmitButton > button:hover {
            transform: translateY(-1.5px);
            box-shadow: 0 4px 12px rgba(111, 78, 55, 0.22);
        }
        .stButton > button:active, .stFormSubmitButton > button:active {
            transform: scale(0.98);
        }
        .stTabs [data-baseweb="tab"] {
            transition: background-color 0.18s ease, color 0.18s ease, transform 0.15s ease;
        }
        .kpi-card::before {
            content: "";
            position: absolute;
            top: 0; left: 0; right: 0;
            height: 4px;
            background: linear-gradient(90deg, #6F4E37, #D49A6A);
        }
        .kpi-label {
            font-size: 0.8rem;
            font-weight: 700;
            color: #7A685C;
            margin-bottom: 0.4rem;
            display: flex;
            align-items: center;
            gap: 0.45rem;
        }
        .kpi-value {
            font-size: 1.6rem;
            font-weight: 800;
            color: #2C1E16;
            line-height: 1.2;
            letter-spacing: -0.02em;
        }
        .kpi-sub {
            font-size: 0.75rem;
            color: #9A877A;
            margin-top: 0.3rem;
        }
        .card-title {
            font-size: 1.05rem;
            font-weight: 700;
            color: #2C1E16;
            margin-bottom: 0.15rem;
        }
        .card-sub {
            font-size: 0.82rem;
            color: #7A685C;
            margin-bottom: 0.75rem;
        }

        /* Metric cards (untuk tab Deep Learning) */
        [data-testid="stMetric"] {
            background-color: #FFFFFF;
            border: 1px solid #EAE0D5;
            padding: 1rem 1.15rem;
            border-radius: 14px;
            box-shadow: 0 2px 10px rgba(44, 30, 22, 0.04);
        }
        [data-testid="stMetricLabel"] {
            font-size: 0.82rem !important;
            color: #6B5B53 !important;
            font-weight: 600 !important;
        }
        [data-testid="stMetricValue"] {
            font-size: 1.45rem !important;
            font-weight: 800 !important;
            color: #2C1E16 !important;
        }

        /* Tabs */
        .stTabs [data-baseweb="tab-list"] {
            gap: 0.35rem;
            background-color: #EFE8DF;
            padding: 0.38rem;
            border-radius: 14px;
            border: 1px solid #E5D9CC;
            margin-bottom: 0.5rem;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius: 10px;
            padding: 0.52rem 1.1rem;
            font-weight: 600;
            color: #5C493E;
        }
        .stTabs [aria-selected="true"] {
            background-color: #FFFFFF !important;
            color: #6F4E37 !important;
            box-shadow: 0 2px 8px rgba(44, 30, 22, 0.08);
        }

        /* Sidebar */
        [data-testid="stSidebar"] {
            background-color: #F4EDE4;
            border-right: 1px solid #E2D4C3;
        }
        [data-testid="stSidebar"] .block-container,
        [data-testid="stSidebarUserContent"] {
            padding-top: 1.1rem !important;
        }
        .sidebar-brand {
            background: linear-gradient(135deg, #2D1B14 0%, #6F4E37 100%);
            color: #FFFDF9;
            padding: 1rem 1.1rem;
            border-radius: 14px;
            margin-bottom: 0.85rem;
            box-shadow: 0 6px 16px rgba(45, 27, 20, 0.14);
        }
        .sidebar-brand-title {
            font-size: 1.05rem;
            font-weight: 800;
            margin: 0;
            letter-spacing: -0.01em;
        }
        .sidebar-brand-sub {
            font-size: 0.76rem;
            opacity: 0.85;
            margin-top: 0.2rem;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"] {
            background-color: #FFFFFF;
            border: 1px solid #E6DCD1 !important;
            border-radius: 14px !important;
            padding: 0.45rem 0.55rem !important;
            box-shadow: 0 3px 10px rgba(44, 30, 22, 0.035);
            transition: box-shadow 0.2s ease;
        }
        [data-testid="stSidebar"] [data-testid="stVerticalBlockBorderWrapper"]:hover {
            box-shadow: 0 5px 14px rgba(44, 30, 22, 0.07);
        }
        [data-testid="stSidebar"] [data-testid="stBaseButton-pills"] {
            background-color: #FAF6F1 !important;
            border: 1px solid #E2D4C3 !important;
            color: #5C493E !important;
            border-radius: 999px !important;
            font-weight: 600 !important;
            font-size: 0.8rem !important;
            transition: all 0.15s ease !important;
        }
        [data-testid="stSidebar"] [data-testid="stBaseButton-pillsActive"] {
            background-color: #6F4E37 !important;
            border: 1px solid #6F4E37 !important;
            color: #FFFFFF !important;
            border-radius: 999px !important;
            font-weight: 700 !important;
            font-size: 0.8rem !important;
            box-shadow: 0 2px 8px rgba(111, 78, 55, 0.25) !important;
        }
        .sidebar-title {
            font-size: 1.08rem;
            font-weight: 800;
            color: #3E2723;
            margin-bottom: 0.1rem;
        }
        .sidebar-caption {
            font-size: 0.78rem;
            color: #7A685C;
            margin-bottom: 0.6rem;
        }
        .filter-label {
            font-size: 0.8rem;
            font-weight: 700;
            color: #5C493E;
            text-transform: uppercase;
            letter-spacing: 0.04em;
            margin-bottom: 0.4rem;
        }
        .sidebar-summary {
            background: linear-gradient(135deg, #FAF6F1, #F2E8DC);
            border: 1px solid #E2D4C3;
            border-radius: 10px;
            padding: 0.65rem 0.85rem;
            font-size: 0.8rem;
            color: #4A3728;
            text-align: center;
            margin-top: 0.25rem;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def rp(x: float) -> str:
    return "Rp" + f"{x:,.0f}".replace(",", ".")


# --------------------------------------------------------------------------
# Parser Excel & Sumber Data
# --------------------------------------------------------------------------
def _cari_sheet(xls: pd.ExcelFile, nama: str) -> str:
    for sh in xls.sheet_names:
        if sh.strip().lower() == nama.lower():
            return sh
    raise ValueError(f'Sheet "{nama}" tidak ditemukan. Sheet yang ada: {", ".join(xls.sheet_names)}')


def parse_penjualan(raw: pd.DataFrame) -> pd.DataFrame:
    """Ubah sheet berblok (Batch N / header produk / Minggu N / TOTAL) menjadi tabel datar."""
    rows, batch, produk = [], None, None
    for _, r in raw.iterrows():
        a = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ""
        isi = r.iloc[1:]
        m_batch = re.match(r"(?i)^batch\s*(\d+)", a)
        m_minggu = re.match(r"(?i)^minggu\s*(\d+)", a)
        if m_batch:
            batch, produk = int(m_batch.group(1)), None
        elif a == "" and isi.notna().any() and all(isinstance(v, str) for v in isi.dropna()):
            produk = [str(v).strip() if pd.notna(v) else None for v in isi]
        elif m_minggu and batch is not None and produk:
            d = {"Batch": batch, "Minggu": int(m_minggu.group(1))}
            for nama, v in zip(produk, isi):
                if nama:
                    d[nama] = pd.to_numeric(v, errors="coerce")
            rows.append(d)
    if not rows:
        raise ValueError(
            'Tidak ada data mingguan di sheet "Penjualan" '
            '(format yang dibaca: "Batch N", baris header produk, lalu "Minggu N").'
        )
    df = pd.DataFrame(rows)
    kolom_produk = [c for c in df.columns if c not in ("Batch", "Minggu")]
    df[kolom_produk] = df[kolom_produk].fillna(0)
    return df[["Batch", "Minggu"] + kolom_produk]


def parse_menu(raw: pd.DataFrame, produk: list) -> pd.DataFrame:
    """Baca sheet Menu dan samakan namanya dengan nama produk di sheet Penjualan."""
    idx = next(
        (
            i
            for i, r in raw.iterrows()
            if {"nama", "hpp"} <= {str(v).strip().lower() for v in r.dropna()}
        ),
        None,
    )
    if idx is None:
        raise ValueError('Sheet "Menu" harus punya kolom: Nama, HPP, Harga Jual.')
    menu = raw.iloc[idx + 1 :].copy()
    menu.columns = [str(v).strip() for v in raw.iloc[idx]]
    for k in ("Nama", "HPP", "Harga Jual"):
        if k not in menu.columns:
            raise ValueError(f'Sheet "Menu" kekurangan kolom "{k}".')
    menu = menu.dropna(subset=["Nama"])
    menu["_n"] = menu["Nama"].astype(str).str.strip().str.lower()

    hasil = []
    for p in produk:
        pl = p.lower()
        cocok = menu[menu["_n"] == pl]
        if cocok.empty:
            cocok = menu[menu["_n"].apply(lambda n: n in pl or pl in n)]
        if cocok.empty:
            raise ValueError(f'Produk "{p}" tidak ada di sheet Menu.')
        r = cocok.iloc[0]
        hasil.append({"Produk": p, "HPP": float(r["HPP"]), "Harga Jual": float(r["Harga Jual"])})
    return pd.DataFrame(hasil)


@st.cache_data(show_spinner=False)
def baca_excel(path: str, mtime: float):
    """Baca file Excel lokal."""
    xls = pd.ExcelFile(path)
    sales = parse_penjualan(pd.read_excel(xls, _cari_sheet(xls, "Penjualan"), header=None))
    produk = [c for c in sales.columns if c not in ("Batch", "Minggu")]
    menu = parse_menu(pd.read_excel(xls, _cari_sheet(xls, "Menu"), header=None), produk)
    return sales, menu


@st.cache_data(ttl=60, show_spinner=False)
def baca_gsheet_cached(sheet_url: str, sa_client_email: str):
    """Baca data dari Google Sheets (Apps Script atau Service Account) dengan cache 60 detik."""
    if is_apps_script_url(sheet_url):
        return read_apps_script_tables(sheet_url, parse_penjualan, parse_menu)
    sa_info, _ = get_gsheet_config(st.secrets)
    sh = connect_spreadsheet(sa_info, sheet_url)
    return read_gsheet_tables(sh, parse_penjualan, parse_menu)


def simpan_tabel_gsheet(sa_info, sheet_url: str, sales_df: pd.DataFrame, menu_df: pd.DataFrame):
    """Simpan tabel Penjualan & Menu ke Google Sheets via Apps Script atau Service Account."""
    if is_apps_script_url(sheet_url):
        write_apps_script_tables(sheet_url, sales_df, menu_df)
    else:
        sh = connect_spreadsheet(sa_info, sheet_url)
        write_full_tables(sh, sales_df, menu_df)


def muat_data():
    """Muat data dari Google Sheets jika terkonfigurasi, atau fallback ke Excel lokal."""
    sa_info, sheet_url = get_gsheet_config(st.secrets)
    gsheet_err = None

    if sheet_url and (is_apps_script_url(sheet_url) or sa_info):
        try:
            email = sa_info.get("client_email", "") if sa_info else "apps_script"
            sales, menu = baca_gsheet_cached(sheet_url, email)
            return sales, menu, "gsheet", None
        except Exception as e:
            gsheet_err = str(e)

    if not FILE_DATA.exists():
        st.error(
            f"Google Sheets belum terhubung dan file **{FILE_DATA.name}** tidak ditemukan di `{FILE_DATA}`."
        )
        if gsheet_err:
            st.error(f"Detail error Google Sheets: {gsheet_err}")
        st.stop()

    try:
        sales, menu = baca_excel(str(FILE_DATA), FILE_DATA.stat().st_mtime)
        return sales, menu, "excel", gsheet_err
    except Exception as e:
        st.error(f"Gagal membaca **{FILE_DATA.name}**: {e}")
        st.stop()


# --------------------------------------------------------------------------
# Deep Learning (TensorFlow / Keras)
# --------------------------------------------------------------------------
def make_xy(vals: np.ndarray, lag: int, scale: float):
    """Ubah deret waktu (T x P) menjadi sampel supervised."""
    T, P = vals.shape
    eye = np.eye(P)
    X, y = [], []
    for i in range(P):
        for t in range(lag, T):
            X.append(np.concatenate([vals[t - lag : t, i] / scale, eye[i]]))
            y.append(vals[t, i] / scale)
    return np.array(X, "float32"), np.array(y, "float32")


def train_ensemble(vals, lag, units, epochs, n_models, scale):
    """Latih beberapa jaringan saraf (MLP) dan kembalikan ensemble-nya."""
    import tensorflow as tf
    from tensorflow import keras
    from tensorflow.keras import layers, regularizers

    X, y = make_xy(vals, lag, scale)
    models, loss_curve = [], None
    for k in range(n_models):
        tf.keras.utils.set_random_seed(42 + k)
        model = keras.Sequential(
            [
                layers.Input(shape=(X.shape[1],)),
                layers.Dense(
                    units, activation="relu", kernel_regularizer=regularizers.l2(1e-3)
                ),
                layers.Dropout(0.1),
                layers.Dense(max(units // 2, 4), activation="relu"),
                layers.Dense(1),
            ]
        )
        model.compile(optimizer=keras.optimizers.Adam(0.01), loss="mse")
        hist = model.fit(X, y, epochs=epochs, batch_size=8, verbose=0)
        models.append(model)
        if k == 0:
            loss_curve = hist.history["loss"]
    return models, loss_curve, len(X)


def predict_next(models, window, scale):
    """Prediksi 1 langkah ke depan untuk semua produk. window: (lag x P)."""
    P = window.shape[1]
    eye = np.eye(P)
    X = np.array(
        [np.concatenate([window[:, i] / scale, eye[i]]) for i in range(P)], "float32"
    )
    preds = np.mean([m(X, training=False).numpy().ravel() for m in models], axis=0)
    return np.clip(preds * scale, 0, None)


@st.cache_data(show_spinner=False)
def run_model(wide: pd.DataFrame, lag: int, units: int, epochs: int, n_models: int, horizon: int):
    vals = wide.values.astype(float)
    scale = float(vals.max()) or 1.0

    train_vals = vals[:-1]
    models_v, _, _ = train_ensemble(train_vals, lag, units, epochs, n_models, scale)
    val_pred = predict_next(models_v, train_vals[-lag:], scale)
    val_true = vals[-1]
    val_naive = train_vals[-1]

    models, loss_curve, n_samples = train_ensemble(vals, lag, units, epochs, n_models, scale)
    window = vals[-lag:].copy()
    forecast = []
    for _ in range(horizon):
        nxt = predict_next(models, window, scale)
        forecast.append(nxt)
        window = np.vstack([window[1:], nxt])
    return {
        "produk": list(wide.columns),
        "val_true": val_true,
        "val_pred": val_pred,
        "val_naive": val_naive,
        "forecast": np.array(forecast),
        "loss": loss_curve,
        "n_samples": n_samples,
    }


# --------------------------------------------------------------------------
# Render Halaman Utama
# --------------------------------------------------------------------------
inject_custom_css()
df_sales, df_menu, data_source, gsheet_error = muat_data()
PRODUK = [c for c in df_sales.columns if c not in ("Batch", "Minggu")]

status_html = (
    '<span class="status-pill status-gsheet">🟢 Google Sheets Terhubung</span>'
    if data_source == "gsheet"
    else '<span class="status-pill status-excel">🟡 Mode Excel Lokal (Fallback)</span>'
)

st.markdown(
    f"""
    <div class="hero-card">
        <div>
            <div class="hero-badge">☕ KELOMPOK 04 · ANALISIS DATA & DEEP LEARNING</div>
            <h1 class="hero-title">Dashboard Analitik Penjualan Kafe</h1>
            <div class="hero-sub">
                Input data otomatis ke Google Sheets · Visualisasi interaktif · Prediksi TensorFlow/Keras
            </div>
        </div>
        <div>{status_html}</div>
    </div>
    """,
    unsafe_allow_html=True,
)

if gsheet_error:
    st.warning(
        f"Koneksi Google Sheets bermasalah (`{gsheet_error}`). Saat ini menampilkan data cadangan dari Excel lokal."
    )

tab_ringkas, tab_tren, tab_komposisi, tab_dl, tab_data = st.tabs(
    [
        "📊 Ringkasan",
        "📈 Tren",
        "🧩 Komposisi",
        "🧠 Prediksi Deep Learning",
        "📝 Input & Data",
    ]
)

# ---- Tab Input & Data ----------------------------------------------------
with tab_data:
    sa_info, sheet_url = get_gsheet_config(st.secrets)
    gsheet_ready = bool(sheet_url and (is_apps_script_url(sheet_url) or sa_info))

    if save_msg := st.session_state.pop("save_anim", None):
        st.balloons()
        st.toast(save_msg, icon="✅")
        st.success(save_msg)

    st.subheader("➕ Input Data Penjualan Mingguan")
    if gsheet_ready:
        st.caption("Data yang diinput di bawah ini akan langsung tersimpan ke **Google Sheets** Anda.")
    else:
        st.info(
            "Google Sheets belum terhubung. Hubungkan di bagian **Pengaturan Google Sheets** di bawah "
            "agar form ini bisa langsung menyimpan data ke Google Sheets."
        )

    last_batch = int(df_sales["Batch"].max()) if not df_sales.empty else 1
    last_minggu_in_batch = (
        int(df_sales.loc[df_sales["Batch"] == last_batch, "Minggu"].max())
        if not df_sales.empty
        else 0
    )

    with st.form("form_input_penjualan", clear_on_submit=False):
        c_b, c_m = st.columns(2)
        in_batch = c_b.number_input("Batch ke-", min_value=1, step=1, value=last_batch)
        in_minggu = c_m.number_input(
            "Minggu ke- (dalam Batch)", min_value=1, step=1, value=last_minggu_in_batch + 1
        )

        st.markdown("**Jumlah Terjual per Produk (cup):**")
        qty_inputs = {}
        cols = st.columns(min(len(PRODUK), 4))
        for idx, p in enumerate(PRODUK):
            qty_inputs[p] = cols[idx % len(cols)].number_input(
                p, min_value=0, step=1, value=0, key=f"in_qty_{p}"
            )

        submitted = st.form_submit_button(
            "💾 Simpan",
            type="primary",
            disabled=not gsheet_ready,
            use_container_width=True,
        )
        if submitted and gsheet_ready:
            with st.spinner("Menyimpan data ke Google Sheets..."):
                try:
                    updated_sales = build_upserted_sales_df(
                        batch=int(in_batch),
                        minggu=int(in_minggu),
                        qty_map=qty_inputs,
                        produk_list=PRODUK,
                        current_sales_df=df_sales,
                    )
                    simpan_tabel_gsheet(sa_info, sheet_url, updated_sales, df_menu)
                    st.cache_data.clear()
                    st.session_state["save_anim"] = (
                        f"Berhasil menyimpan Batch {int(in_batch)} - Minggu {int(in_minggu)} ke Google Sheets!"
                    )
                    st.rerun()
                except Exception as e:
                    st.error(f"Gagal menyimpan ke Google Sheets: {e}")

    st.divider()

    # Editor Tabel Penjualan & Menu
    st.subheader("🗂️ Tabel Data Penjualan & Menu")
    st.caption("Anda bisa mengedit langsung tabel di bawah ini, lalu simpan perubahannya.")
    sales_in = st.data_editor(
        df_sales, num_rows="dynamic", key="sales", use_container_width=True, hide_index=True
    )
    st.markdown("**Daftar Menu & HPP**")
    menu_in = st.data_editor(
        df_menu, disabled=["Produk"], key="menu", use_container_width=True, hide_index=True
    )

    if gsheet_ready and st.button("💾 Simpan Perubahan Tabel"):
        with st.spinner("Memperbarui Google Sheets..."):
            try:
                clean_s = sales_in.dropna(subset=["Batch", "Minggu"]).copy()
                clean_s[PRODUK] = clean_s[PRODUK].fillna(0)
                simpan_tabel_gsheet(sa_info, sheet_url, clean_s, menu_in.fillna(0))
                st.cache_data.clear()
                st.session_state["save_anim"] = "Semua perubahan tabel berhasil disimpan ke Google Sheets!"
                st.rerun()
            except Exception as e:
                st.error(f"Gagal memperbarui Google Sheets: {e}")

# ---- Olah Data -----------------------------------------------------------
sales = sales_in.dropna(subset=["Batch", "Minggu"]).copy()
sales[PRODUK] = sales[PRODUK].fillna(0)
sales = sales.astype({"Batch": int, "Minggu": int}).sort_values(["Batch", "Minggu"])
sales = sales.reset_index(drop=True)
sales["Minggu ke"] = np.arange(1, len(sales) + 1)
menu = menu_in.fillna(0)

if sales.empty:
    st.warning("Data penjualan kosong. Isi data di tab Input & Data.")
    st.stop()

long = sales.melt(
    id_vars=["Batch", "Minggu", "Minggu ke"], value_vars=PRODUK, var_name="Produk", value_name="Qty"
).merge(menu, on="Produk")
long["Omzet"] = long["Qty"] * long["Harga Jual"]
long["Laba"] = long["Qty"] * (long["Harga Jual"] - long["HPP"])
long["Periode"] = "B" + long["Batch"].astype(str) + "-M" + long["Minggu"].astype(str)
periode_order = list(sales["Batch"].astype(str).radd("B") + "-M" + sales["Minggu"].astype(str))

# ---- Sidebar Filter ------------------------------------------------------
all_batches = sorted(long["Batch"].unique())
with st.sidebar:
    st.markdown(
        """
        <div class="sidebar-brand">
            <div class="sidebar-brand-title">☕ Kontrol Dashboard</div>
            <div class="sidebar-brand-sub">Filter metrik, batch &amp; menu kopi</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    sb1, sb2 = st.columns(2)
    if sb1.button("🔄 Refresh", use_container_width=True, type="primary"):
        st.cache_data.clear()
        st.rerun()
    if sb2.button("↺ Reset", use_container_width=True):
        st.session_state["f_metrik"] = "Qty"
        st.session_state["f_batch"] = all_batches
        st.session_state["f_produk"] = PRODUK
        st.rerun()

    with st.container(border=True):
        st.markdown('<div class="filter-label">📊 Metrik Utama</div>', unsafe_allow_html=True)
        metrik = st.pills(
            "Metrik Utama",
            list(METRIK),
            selection_mode="single",
            default="Qty",
            key="f_metrik",
            format_func=lambda m: {"Qty": "Qty (cup)", "Omzet": "Omzet", "Laba": "Laba"}[m],
            label_visibility="collapsed",
        )
        if not metrik:
            metrik = "Qty"

    with st.container(border=True):
        st.markdown('<div class="filter-label">📦 Filter Batch</div>', unsafe_allow_html=True)
        batch_sel = st.pills(
            "Batch",
            all_batches,
            selection_mode="multi",
            default=all_batches,
            key="f_batch",
            format_func=lambda b: f"Batch {b}",
            label_visibility="collapsed",
        )
        if not batch_sel:
            batch_sel = all_batches

    with st.container(border=True):
        st.markdown('<div class="filter-label">🥤 Filter Produk</div>', unsafe_allow_html=True)
        produk_sel = st.pills(
            "Produk",
            PRODUK,
            selection_mode="multi",
            default=PRODUK,
            key="f_produk",
            label_visibility="collapsed",
        )
        if not produk_sel:
            produk_sel = PRODUK

    st.markdown(
        f"""
        <div class="sidebar-summary">
            <b>{len(batch_sel)}/{len(all_batches)}</b> Batch · <b>{len(produk_sel)}/{len(PRODUK)}</b> Produk
        </div>
        """,
        unsafe_allow_html=True,
    )

f = long[long["Batch"].isin(batch_sel) & long["Produk"].isin(produk_sel)]
if f.empty:
    st.warning("Tidak ada data untuk filter yang dipilih.")
    st.stop()

# ---- Ringkasan -----------------------------------------------------------
with tab_ringkas:
    best = f.groupby("Produk")["Qty"].sum().idxmax()
    best_qty = int(f.groupby("Produk")["Qty"].sum().max())
    total_qty = f"{int(f['Qty'].sum()):,} cup".replace(",", ".")
    total_omzet = rp(f["Omzet"].sum())
    total_laba = rp(f["Laba"].sum())
    margin_pct = (f["Laba"].sum() / f["Omzet"].sum() * 100) if f["Omzet"].sum() else 0

    st.markdown(
        f"""
        <div class="kpi-grid">
            <div class="kpi-card">
                <div class="kpi-label">☕ Total Terjual</div>
                <div class="kpi-value">{total_qty}</div>
                <div class="kpi-sub">Dari {len(f['Periode'].unique())} periode aktif</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">💰 Total Omzet</div>
                <div class="kpi-value">{total_omzet}</div>
                <div class="kpi-sub">Pendapatan kotor penjualan</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">📈 Total Laba Kotor</div>
                <div class="kpi-value">{total_laba}</div>
                <div class="kpi-sub">Margin rata-rata {margin_pct:.1f}%</div>
            </div>
            <div class="kpi-card">
                <div class="kpi-label">🏆 Produk Terlaris</div>
                <div class="kpi-value">{best}</div>
                <div class="kpi-sub">{best_qty:,} cup terjual</div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.markdown("")
    col_left, col_right = st.columns(2)
    with col_left:
        with st.container(border=True):
            st.markdown('<div class="card-title">📦 Agregasi per Batch</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-sub">Ringkasan total cup, omzet, dan laba setiap batch</div>', unsafe_allow_html=True)
            agg = f.groupby("Batch")[["Qty", "Omzet", "Laba"]].sum().reset_index()
            st.dataframe(
                agg.style.format({"Omzet": rp, "Laba": rp}),
                use_container_width=True,
                hide_index=True,
            )
    with col_right:
        with st.container(border=True):
            st.markdown('<div class="card-title">🥤 Agregasi per Produk</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-sub">Performa penjualan diurutkan dari yang terlaris</div>', unsafe_allow_html=True)
            aggp = (
                f.groupby("Produk")[["Qty", "Omzet", "Laba"]]
                .sum()
                .sort_values("Qty", ascending=False)
                .reset_index()
            )
            st.dataframe(
                aggp.style.format({"Omzet": rp, "Laba": rp}),
                use_container_width=True,
                hide_index=True,
            )

    with st.container(border=True):
        st.markdown(f'<div class="card-title">📊 Peringkat Produk ({METRIK[metrik]})</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-sub">Perbandingan cepat kontribusi masing-masing menu</div>', unsafe_allow_html=True)
        bar_rank = aggp.sort_values(metrik, ascending=True)
        fig_rank = px.bar(
            bar_rank,
            x=metrik,
            y="Produk",
            orientation="h",
            color="Produk",
            color_discrete_sequence=COFFEE_PALETTE,
            text_auto=True,
            template="plotly_white",
        )
        fig_rank.update_layout(
            showlegend=False,
            height=260,
            margin=dict(l=10, r=20, t=10, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            xaxis=dict(gridcolor="#F0E9E0", title=""),
            yaxis=dict(title=""),
        )
        st.plotly_chart(fig_rank, use_container_width=True)

# ---- Tren ----------------------------------------------------------------
with tab_tren:
    with st.container(border=True):
        st.markdown(f'<div class="card-title">📈 Tren Mingguan — {METRIK[metrik]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-sub">Pergerakan penjualan dari minggu ke minggu tiap produk</div>', unsafe_allow_html=True)
        trend = (
            f.groupby(["Minggu ke", "Periode", "Produk"], as_index=False)[metrik]
            .sum()
            .sort_values("Minggu ke")
        )
        fig = px.line(
            trend,
            x="Periode",
            y=metrik,
            color="Produk",
            markers=True,
            line_shape="spline",
            category_orders={"Periode": periode_order},
            color_discrete_sequence=COFFEE_PALETTE,
            template="plotly_white",
        )
        fig.update_traces(line=dict(width=3), marker=dict(size=7))
        fig.update_layout(
            hovermode="x unified",
            legend_title_text="",
            margin=dict(l=10, r=10, t=15, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#F0E9E0"),
            xaxis=dict(showgrid=False),
        )
        st.plotly_chart(fig, use_container_width=True)

    with st.container(border=True):
        st.markdown(f'<div class="card-title">📊 Perbandingan Antar Batch — {METRIK[metrik]}</div>', unsafe_allow_html=True)
        st.markdown('<div class="card-sub">Performa masing-masing produk di setiap batch</div>', unsafe_allow_html=True)
        bb = f.groupby(["Batch", "Produk"], as_index=False)[metrik].sum()
        bb["Batch"] = "Batch " + bb["Batch"].astype(str)
        fig2 = px.bar(
            bb,
            x="Batch",
            y=metrik,
            color="Produk",
            barmode="group",
            color_discrete_sequence=COFFEE_PALETTE,
            template="plotly_white",
            text_auto=True,
        )
        fig2.update_layout(
            legend_title_text="",
            margin=dict(l=10, r=10, t=15, b=10),
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
            yaxis=dict(gridcolor="#F0E9E0"),
        )
        st.plotly_chart(fig2, use_container_width=True)

# ---- Komposisi -----------------------------------------------------------
with tab_komposisi:
    col_a, col_b = st.columns(2)
    with col_a:
        with st.container(border=True):
            st.markdown(f'<div class="card-title">🍩 Kontribusi Produk — {METRIK[metrik]}</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-sub">Persentase porsi penjualan per menu</div>', unsafe_allow_html=True)
            share = f.groupby("Produk", as_index=False)[metrik].sum()
            figp = px.pie(
                share,
                names="Produk",
                values=metrik,
                hole=0.55,
                color_discrete_sequence=COFFEE_PALETTE,
                template="plotly_white",
            )
            figp.update_traces(textposition="inside", textinfo="percent+label")
            figp.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=10, r=10, t=15, b=10),
                showlegend=False,
            )
            st.plotly_chart(figp, use_container_width=True)
    with col_b:
        with st.container(border=True):
            st.markdown('<div class="card-title">🔥 Heatmap Qty (Periode × Produk)</div>', unsafe_allow_html=True)
            st.markdown('<div class="card-sub">Intensitas jumlah cup terjual di tiap periode</div>', unsafe_allow_html=True)
            pv = f.pivot_table(index="Periode", columns="Produk", values="Qty", aggfunc="sum")
            pv = pv.reindex([p for p in periode_order if p in pv.index])
            figh = px.imshow(
                pv,
                text_auto=True,
                aspect="auto",
                color_continuous_scale=["#FAF3E8", "#C68B59", "#6F4E37"],
            )
            figh.update_layout(
                paper_bgcolor="rgba(0,0,0,0)",
                margin=dict(l=10, r=10, t=15, b=10),
            )
            st.plotly_chart(figh, use_container_width=True)

# ---- Deep Learning -------------------------------------------------------
with tab_dl:
    st.subheader("Prediksi Penjualan dengan Deep Learning")
    st.markdown(
        "Model: **jaringan saraf tiruan berlapis (MLP) dengan TensorFlow/Keras**, "
        "dilatih dari penjualan beberapa minggu terakhir + identitas produk, "
        "lalu dibuat *ensemble* beberapa model agar hasil lebih stabil."
    )
    with st.expander("⚙️ Pengaturan Parameter Model", expanded=True):
        p1, p2, p3, p4, p5 = st.columns(5)
        lag = p1.slider("Jendela (minggu)", 1, 3, 2)
        units = p2.slider("Neuron/lapisan", 8, 64, 16, step=8)
        epochs = p3.slider("Epoch", 50, 500, 200, step=50)
        n_models = p4.slider("Jumlah model", 1, 10, 5)
        horizon = p5.slider("Prediksi ke depan (minggu)", 1, 4, 2)

    wide = sales.set_index("Minggu ke")[PRODUK]
    if len(wide) < lag + 2:
        st.warning("Data belum cukup untuk jendela ini. Tambah data atau kecilkan jendela.")
    elif st.button("🚀 Latih model & prediksi", type="primary"):
        with st.status("🧠 Melatih ensemble jaringan saraf...", expanded=True) as status:
            try:
                st.write("Menyiapkan data deret waktu & normalisasi fitur...")
                prog = st.progress(25)
                st.write(f"Melatih {n_models} model MLP ({epochs} epoch) dengan TensorFlow/Keras...")
                prog.progress(60)
                st.session_state["dl"] = run_model(wide, lag, units, epochs, n_models, horizon)
                prog.progress(100)
                status.update(label="✅ Model selesai dilatih & prediksi siap!", state="complete", expanded=False)
                st.toast("Prediksi Deep Learning berhasil diperbarui!", icon="🧠")
            except ImportError:
                status.update(label="❌ TensorFlow belum terpasang", state="error")
                st.error("TensorFlow belum terpasang. Jalankan: pip install tensorflow")

    res = st.session_state.get("dl")
    if res:
        prods = res["produk"]
        mae_dl = float(np.mean(np.abs(res["val_true"] - res["val_pred"])))
        mae_nv = float(np.mean(np.abs(res["val_true"] - res["val_naive"])))
        m1, m2, m3 = st.columns(3)
        m1.metric("MAE Deep Learning", f"{mae_dl:.2f} cup")
        m2.metric("MAE Baseline (minggu lalu)", f"{mae_nv:.2f} cup")
        m3.metric("Sampel latih", res["n_samples"])
        st.caption(
            "MAE dihitung dengan menyembunyikan minggu terakhir lalu menebaknya. Semakin kecil semakin baik."
        )

        last_wk = int(wide.index.max())
        figf = go.Figure()
        for i, p in enumerate(prods):
            if p not in produk_sel:
                continue
            color = COFFEE_PALETTE[i % len(COFFEE_PALETTE)]
            figf.add_trace(
                go.Scatter(
                    x=wide.index,
                    y=wide[p],
                    name=p,
                    mode="lines+markers",
                    line=dict(color=color),
                )
            )
            fx = [last_wk] + [last_wk + k + 1 for k in range(len(res["forecast"]))]
            fy = [wide[p].iloc[-1]] + list(res["forecast"][:, i])
            figf.add_trace(
                go.Scatter(
                    x=fx,
                    y=fy,
                    name=f"{p} (prediksi)",
                    mode="lines+markers",
                    line=dict(color=color, dash="dash"),
                    marker=dict(symbol="diamond"),
                )
            )
        figf.update_layout(
            template="plotly_white",
            hovermode="x unified",
            xaxis_title="Minggu ke-",
            yaxis_title="Qty (cup)",
            paper_bgcolor="rgba(0,0,0,0)",
            plot_bgcolor="rgba(0,0,0,0)",
        )
        st.plotly_chart(figf, use_container_width=True)

        harga = menu.set_index("Produk")
        rows = []
        for k, arr in enumerate(res["forecast"]):
            for i, p in enumerate(prods):
                q = float(arr[i])
                rows.append(
                    {
                        "Minggu ke-": last_wk + k + 1,
                        "Produk": p,
                        "Prediksi Qty": round(q, 1),
                        "Est. Omzet": q * harga.loc[p, "Harga Jual"],
                        "Est. Laba": q * (harga.loc[p, "Harga Jual"] - harga.loc[p, "HPP"]),
                    }
                )
        st.dataframe(
            pd.DataFrame(rows).style.format({"Est. Omzet": rp, "Est. Laba": rp}),
            use_container_width=True,
            hide_index=True,
        )

        cv, cl = st.columns(2)
        with cv:
            st.markdown("**Validasi minggu terakhir**")
            st.dataframe(
                pd.DataFrame(
                    {
                        "Produk": prods,
                        "Aktual": res["val_true"],
                        "Prediksi DL": res["val_pred"].round(1),
                        "Baseline": res["val_naive"],
                    }
                ),
                use_container_width=True,
                hide_index=True,
            )
        with cl:
            st.markdown("**Kurva loss pelatihan**")
            figl = px.line(
                y=res["loss"],
                labels={"x": "Epoch", "y": "MSE"},
                color_discrete_sequence=["#6F4E37"],
                template="plotly_white",
            )
            figl.update_layout(paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(figl, use_container_width=True)

        st.info(
            "Catatan: data hanya beberapa minggu, jadi prediksi ini lebih bersifat "
            "demonstrasi metode. Makin panjang riwayat datanya, makin layak dipercaya."
        )
    else:
        st.info("Atur parameter lalu klik **Latih model & prediksi**.")

with tab_data:
    st.download_button(
        "⬇️ Unduh data (CSV)",
        long.to_csv(index=False).encode(),
        "data_penjualan_long.csv",
        "text/csv",
    )

st.markdown(
    """
    <div style="text-align:center; color:#8C7A6B; font-size:0.8rem; margin-top:2.5rem; padding-top:1rem; border-top:1px solid #EAE0D5;">
        ☕ <b>Dashboard Analitik Penjualan Kafe</b> · Kelompok 04 · Powered by Streamlit &amp; TensorFlow/Keras
    </div>
    """,
    unsafe_allow_html=True,
)
