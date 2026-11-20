"""
Helper integrasi Google Sheets (baca & tulis) untuk Dashboard Penjualan Kafe.
Mendukung 2 mode koneksi:
  1. Google Apps Script Web App URL (PALING SIMPEL - tanpa Google Cloud / JSON)
  2. Service Account JSON (st.secrets / env / file lokal) -> opsional
"""
import json
import os
from pathlib import Path
from typing import Optional, Tuple

import gspread
import pandas as pd
import requests
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

BASE_DIR = Path(__file__).parent
LOCAL_SA_FILE = BASE_DIR / "service_account.json"
LOCAL_CFG_FILE = BASE_DIR / "gsheet_config.json"

APPS_SCRIPT_TEMPLATE = """function doGet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sPenjualan = ss.getSheetByName("Penjualan");
  var sMenu = ss.getSheetByName("Menu");
  return ContentService.createTextOutput(JSON.stringify({
    penjualan: sPenjualan ? sPenjualan.getDataRange().getValues() : [],
    menu: sMenu ? sMenu.getDataRange().getValues() : []
  })).setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  function getOrCreate(name) {
    return ss.getSheetByName(name) || ss.insertSheet(name);
  }

  if (data.action === "write_all") {
    var s1 = getOrCreate("Penjualan");
    s1.clearContents();
    if (data.penjualan && data.penjualan.length > 0) {
      s1.getRange(1, 1, data.penjualan.length, data.penjualan[0].length).setValues(data.penjualan);
    }
    var s2 = getOrCreate("Menu");
    s2.clearContents();
    if (data.menu && data.menu.length > 0) {
      s2.getRange(1, 1, data.menu.length, data.menu[0].length).setValues(data.menu);
    }
  }
  return ContentService.createTextOutput(JSON.stringify({status: "ok"}))
    .setMimeType(ContentService.MimeType.JSON);
}"""


def is_apps_script_url(url: Optional[str]) -> bool:
    return bool(url and "script.google.com" in url)


def get_gsheet_config(st_secrets=None) -> Tuple[Optional[dict], Optional[str]]:
    """Ambil URL Apps Script / Google Sheet & kredensial opsional dari secrets, env, atau file lokal."""
    sa_info = None
    sheet_url = None

    if st_secrets is not None:
        try:
            if "apps_script_url" in st_secrets:
                sheet_url = str(st_secrets["apps_script_url"]).strip()
            elif "gsheet_url" in st_secrets:
                sheet_url = str(st_secrets["gsheet_url"]).strip()
            if "gcp_service_account" in st_secrets:
                sa_info = dict(st_secrets["gcp_service_account"])
        except Exception:
            pass

    if not sheet_url and os.environ.get("APPS_SCRIPT_URL"):
        sheet_url = os.environ["APPS_SCRIPT_URL"].strip()
    if not sheet_url and os.environ.get("GSHEET_URL"):
        sheet_url = os.environ["GSHEET_URL"].strip()
    if sa_info is None and os.environ.get("GCP_SERVICE_ACCOUNT_JSON"):
        try:
            sa_info = json.loads(os.environ["GCP_SERVICE_ACCOUNT_JSON"])
        except Exception:
            pass

    if not sheet_url and LOCAL_CFG_FILE.exists():
        try:
            cfg = json.loads(LOCAL_CFG_FILE.read_text(encoding="utf-8"))
            sheet_url = str(cfg.get("gsheet_url", "")).strip() or None
        except Exception:
            pass
    if sa_info is None and LOCAL_SA_FILE.exists():
        try:
            sa_info = json.loads(LOCAL_SA_FILE.read_text(encoding="utf-8"))
        except Exception:
            pass

    return sa_info, sheet_url


def save_local_config(sa_json_bytes: Optional[bytes], sheet_url: str) -> None:
    """Simpan konfigurasi secara lokal agar tidak perlu input ulang."""
    if sa_json_bytes:
        LOCAL_SA_FILE.write_bytes(sa_json_bytes)
    if sheet_url:
        LOCAL_CFG_FILE.write_text(
            json.dumps({"gsheet_url": sheet_url.strip()}, indent=2),
            encoding="utf-8",
        )


# --------------------------------------------------------------------------
# Mode 1: Google Apps Script (Tanpa Google Cloud / JSON)
# --------------------------------------------------------------------------
def _parse_tables(raw_sales_rows, raw_menu_rows, parse_penjualan_fn, parse_menu_fn):
    raw_sales = pd.DataFrame(raw_sales_rows)
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

    raw_menu = pd.DataFrame(raw_menu_rows)
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


def read_apps_script_tables(script_url: str, parse_penjualan_fn, parse_menu_fn):
    resp = requests.get(script_url, timeout=20)
    resp.raise_for_status()
    payload = resp.json()
    return _parse_tables(
        payload.get("penjualan", []),
        payload.get("menu", []),
        parse_penjualan_fn,
        parse_menu_fn,
    )


def write_apps_script_tables(script_url: str, sales_df: pd.DataFrame, menu_df: pd.DataFrame) -> None:
    clean_sales = sales_df.drop(columns=["Minggu ke"], errors="ignore").copy()
    payload = {
        "action": "write_all",
        "penjualan": [clean_sales.columns.tolist()] + clean_sales.fillna(0).values.tolist(),
        "menu": [menu_df.columns.tolist()] + menu_df.fillna(0).values.tolist(),
    }
    resp = requests.post(script_url, json=payload, timeout=25)
    resp.raise_for_status()


# --------------------------------------------------------------------------
# Mode 2: Service Account JSON (gspread)
# --------------------------------------------------------------------------
def connect_spreadsheet(sa_info: dict, sheet_url: str) -> gspread.Spreadsheet:
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
    ws_sales = sh.worksheet("Penjualan")
    ws_menu = sh.worksheet("Menu")
    return _parse_tables(
        ws_sales.get_all_values(),
        ws_menu.get_all_values(),
        parse_penjualan_fn,
        parse_menu_fn,
    )


def write_full_tables(sh: gspread.Spreadsheet, sales_df: pd.DataFrame, menu_df: pd.DataFrame) -> None:
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


def build_upserted_sales_df(
    batch: int,
    minggu: int,
    qty_map: dict,
    produk_list: list,
    current_sales_df: pd.DataFrame,
) -> pd.DataFrame:
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

    return df.sort_values(["Batch", "Minggu"]).reset_index(drop=True)
