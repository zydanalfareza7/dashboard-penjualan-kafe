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

## 3. Cara Hosting Gratis Online (Streamlit Community Cloud)

> **Catatan:** Hosting gratis seperti InfinityFree tidak bisa menjalankan Python/TensorFlow, sedangkan Hugging Face Spaces untuk compute/Docker kini memerlukan paket PRO. Gunakan **Streamlit Community Cloud** (`share.streamlit.io`) yang gratis dan terhubung langsung dengan GitHub!

1. Buka [share.streamlit.io](https://share.streamlit.io) dan login dengan akun **GitHub**.
2. Klik **Create app** → pilih **Deploy a public app from GitHub** (atau hubungkan repo Private Anda).
3. Pilih repository: `zydanalfareza7/dashboard-penjualan-kafe`, branch `main`, dan Main file path: `app.py`.
4. Klik **Advanced settings → Secrets** dan masukkan konfigurasi Google Sheets Anda:
   ```toml
   gsheet_url = "https://docs.google.com/spreadsheets/d/..."

   [gcp_service_account]
   # salin-tempel isi JSON service account dalam format TOML
   ```
5. Klik **Deploy!** dan tunggu proses instalasi selesai hingga mendapatkan link publik `*.streamlit.app`.
