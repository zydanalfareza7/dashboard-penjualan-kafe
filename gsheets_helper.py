"""
Helper integrasi Google Sheets (baca & tulis) untuk Dashboard Penjualan Kafe.
Mendukung konfigurasi via:
  1. st.secrets ("gcp_service_account" + "gsheet_url") -> Streamlit Cloud / lokal
  2. Environment variable ("GCP_SERVICE_ACCOUNT_JSON" + "GSHEET_URL") -> Hugging Face Spaces
  3. File lokal ("service_account.json" + "gsheet_config.json") -> Setup langsung dari UI
"""
import json
import os
from pathlib import Path
from typing import Optional, Tuple

import gspread
import pandas as pd
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

BASE_DIR = Path(__file__).parent
LOCAL_SA_FILE = BASE_DIR / "service_account.json"
LOCAL_CFG_FILE = BASE_DIR / "gsheet_config.json"


def get_gsheet_config(st_secrets=None) -> Tuple[Optional[dict], Optional[str]]:
    """Ambil kredensial service account & URL Google Sheet dari secrets, env, atau file lokal."""
    sa_info = None
    sheet_url = None

    # 1. Streamlit secrets
    if st_secrets is not None:
        try:
            if "gcp_service_account" in st_secrets:
                sa_info = dict(st_secrets["gcp_service_account"])
            if "gsheet_url" in st_secrets:
                sheet_url = str(st_secrets["gsheet_url"]).strip()
        except Exception:
            pass

    # 2. Environment variables (cocok untuk Hugging Face Spaces / Docker)
    if sa_info is None and os.environ.get("GCP_SERVICE_ACCOUNT_JSON"):
        try:
            sa_info = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])
        except Exception:
            pass
    if not sheet_url and os.environ.get("GSHEET_URL"):
        sheet_url = os.environ["GSHEET_URL"].strip()

    # 3. File lokal di folder project
    if sa_info is None and LOCAL_SA_FILE.exists():
        try:
            sa_info = json.loads(LOCAL_SA_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass
    if not sheet_url and LOCAL_CFG_FILE.exists():
        try:
            cfg = json.loads(LOCAL_CFG_FILE.read_text(encoding="utf-8"))
            sheet_url = str(cfg.get("gsheet_url", "")).strip() or None
        except Exception:
            pass

    return sa_info, sheet_url


def save_local_config(sa_json_bytes: Optional[bytes], sheet_url: str) -> None:
    """Simpan konfigurasi Google Sheets secara lokal agar tidak perlu input ulang."""
    if sa_json_bytes:
        LOCAL_SA_FILE.write_bytes(sa_json_bytes)
    if sheet_url:
        LOCAL_CFG_FILE.write_text(
            json.dumps({"gsheet_url": sheet_url.strip()}, indent=2),
            encoding="utf-8",
        )


def connect_spreadsheet(sa_info: dict, sheet_url: str) -> gspread.Spreadsheet:
    """Buat koneksi gspread ke Google Spreadsheet target."""
    creds = Credentials.from_service_account_info(sa_info, scopes=SCOPES)
    client = gspread.authorize(creds)
    if sheet_url.startswith("http"):
        return client.open_by_url(sheet_url)
    return client.open_by_key(sheet_url)


def _get_or_create_ws(sh: gspread.Spreadsheet, title: str, rows: int = 200, cols: int = 20):
    try:
        return sh.worksheet(title)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=title, rows=rows, cols=cols)


def read_gsheet_tables(sh: gspread.Spreadsheet, parse_penjualan_fn, parse_menu_fn):
    """Baca sheet Penjualan & Menu dari Google Sheets (mendukung format datar maupun format blok)."""
    ws_sales = sh.worksheet("Penjualan")
    raw_sales = pd.DataFrame(ws_sales.get_all_values())
    if raw_sales.empty:
        raise ValueError('Sheet "Penjualan" di Google Sheets masih kosong.')

    first_row = [str(x).strip().lower() for x in raw_sales.iloc[0].tolist()]
    if "batch" in first_row and "minggu" in first_row:
        sales = raw_sales.iloc[1:].copy()
        sales.columns = [str(x).strip() for x in raw_sales.iloc[0].tolist()]
        sales = sales.replace("", pd.NA).dropna(subset=["Batch", "Minggu"])
        produk = [c for c in sales.columns if c not in ("Batch", "Minggu", "Minggu ke")]
        sales["Batch"] = pd.to_numeric(sales["Batch"], errors="coerce")
        sales["Minggu"] = pd.to_numeric(sales["Minggu"], errors="coerce")
        for p in produk:
            sales[p] = pd.to_numeric(sales[p], errors="coerce").fillna(0)
        sales = sales.dropna(subset=["Batch", "Minggu"])[["Batch", "Minggu"] + produk]
    else:
        sales = parse_penjualan_fn(raw_sales)
        produk = [c for c in sales.columns if c not in ("Batch", "Minggu")]

    ws_menu = sh.worksheet("Menu")
    raw_menu = pd.DataFrame(ws_menu.get_all_values())
    if raw_menu.empty:
        raise ValueError('Sheet "Menu" di Google Sheets masih kosong.')

    menu_header = [str(x).strip().lower() for x in raw_menu.iloc[0].tolist()]
    if {"produk", "hpp", "harga jual"} <= set(menu_header):
        menu = raw_menu.iloc[1:].copy()
        menu.columns = [str(x).strip() for x in raw_menu.iloc[0].tolist()]
        menu = menu[menu["Produk"].astype(str).str.strip() != ""].copy()
        menu["HPP"] = pd.to_numeric(menu["HPP"], errors="coerce").fillna(0)
        menu["Harga Jual"] = pd.to_numeric(menu["Harga Jual"], errors="coerce").fillna(0)
        menu = menu[["Produk", "HPP", "Harga Jual"]]
    else:
        menu = parse_menu_fn(raw_menu, produk)

    return sales, menu


def write_full_tables(sh: gspread.Spreadsheet, sales_df: pd.DataFrame, menu_df: pd.DataFrame) -> None:
    """Tulis ulang seluruh tabel Penjualan & Menu dalam format datar yang rapi ke Google Sheets."""
    ws_sales = _get_or_create_ws(sh, "Penjualan")
    clean_sales = sales_df.drop(columns=["Minggu ke"], errors="ignore").copy()
    ws_sales.clear()
    ws_sales.update(
        [clean_sales.columns.tolist()] + clean_sales.fillna(0).values.tolist(),
        value_input_option="USER_ENTERED",
    )

    ws_menu = _get_or_create_ws(sh, "Menu")
    ws_menu.clear()
    ws_menu.update(
        [menu_df.columns.tolist()] + menu_df.fillna(0).values.tolist(),
        value_input_option="USER_ENTERED",
    )


def upsert_sales_row(
    sh: gspread.Spreadsheet,
    batch: int,
    minggu: int,
    qty_map: dict,
    produk_list: list,
    current_sales_df: pd.DataFrame,
    current_menu_df: pd.DataFrame,
) -> pd.DataFrame:
    """Tambah atau perbarui 1 baris data penjualan (Batch + Minggu) lalu simpan ke Google Sheets."""
    df = current_sales_df.drop(columns=["Minggu ke"], errors="ignore").copy()
    new_row = {"Batch": int(batch), "Minggu": int(minggu)}
    for p in produk_list:
        new_row[p] = float(qty_map.get(p, 0))

    mask = (df["Batch"].astype(int) == int(batch)) & (df["Minggu"].astype(int) == int(minggu))
    if mask.any():
        for k, v in new_row.items():
            df.loc[mask, k] = v
    else:
        df = pd.concat([df, pd.DataFrame([new_row])], ignore_index=True)

    df = df.sort_values(["Batch", "Minggu"]).reset_index(drop=True)
    write_full_tables(sh, df, current_menu_df)
    return df
