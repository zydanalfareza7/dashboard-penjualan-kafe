---
title: Dashboard Analitik Penjualan Kafe
emoji: ☕
colorFrom: yellow
colorTo: gray
sdk: streamlit
app_file: app.py
pinned: false
---

# ☕ Dashboard Analitik Penjualan Kafe (Streamlit + Deep Learning)

Dashboard interaktif untuk menganalisis penjualan kafe, input data langsung ke **Google Sheets**, dan memprediksi penjualan menggunakan **Deep Learning (TensorFlow / Keras)**.

---

## 1. Cara Menjalankan di Laptop (Lokal)

> **Penting:** Gunakan **Python 3.12** (atau 3.10–3.12), karena TensorFlow belum mendukung Python 3.13/3.14.

1. **Clone atau download** repository ini, lalu buka terminal/PowerShell di dalam foldernya.
2. **Buat virtual environment & install library:**
   ```bash
   py -3.12 -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
   *(Di Mac/Linux gunakan `python3 -m venv .venv` dan `source .venv/bin/activate`)*
3. **Jalankan aplikasi:**
   ```bash
   streamlit run app.py
   ```
4. Buka browser ke `http://localhost:8501`.

---

## 2. Cara Menghubungkan Google Sheets (Agar Input Data Otomatis)

Secara default (jika Google Sheets belum disambungkan), aplikasi membaca data cadangan dari `Data_Penjualan_Kopi_1.xlsx`. Supaya input data dari web langsung tersimpan ke Google Sheets:

1. **Buat Google Sheet baru** di Google Drive, lalu salin URL-nya.
2. Buka **Google Cloud Console** (`console.cloud.google.com`):
   - Buat project baru.
   - Aktifkan **Google Sheets API** dan **Google Drive API**.
   - Masuk ke **IAM & Admin → Service Accounts → Create Service Account**.
   - Buka Service Account yang sudah dibuat → tab **Keys → Add Key → Create new key → JSON** (simpan file `.json` ini, jangan di-upload ke GitHub).
3. **Share Google Sheet** kamu ke alamat `client_email` yang ada di dalam file `.json` tersebut (pilih akses **Editor**).
4. Buka web dashboard → masuk ke tab **📝 Input & Data** → buka bagian **☁️ Pengaturan & Koneksi Google Sheets**:
   - Tempel URL Google Sheet dan upload file `.json`.
   - Klik **🔗 Simpan Konfigurasi Google Sheets**.
   - Klik **📤 Salin Data Saat Ini ke Google Sheet** untuk memindahkan data awal dari Excel ke Google Sheet.

---

## 3. Cara Hosting Gratis Online (Hugging Face Spaces)

> **Catatan:** Hosting gratis seperti InfinityFree tidak bisa menjalankan Python/TensorFlow. Gunakan **Hugging Face Spaces** (gratis dengan RAM 16 GB).

1. Buat akun di [huggingface.co](https://huggingface.co), lalu klik **New Space**.
2. Isi nama Space, pilih **SDK: Streamlit**, lalu klik **Create Space**.
3. Upload semua file dari repo ini (kecuali `.venv/`, `service_account.json`, `gsheet_config.json`).
4. Masuk ke menu **Settings → Variables and secrets** di Space tersebut, lalu tambahkan **New secret**:
   - `GSHEET_URL` → isi dengan URL Google Sheet kamu.
   - `GCP_SERVICE_ACCOUNT_JSON` → salin-tempel seluruh isi file `.json` Service Account.
5. Tunggu proses build selesai, dan dashboard siap diakses lewat link publik!
