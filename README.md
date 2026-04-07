# Multibit Quantization & Kalman Automation

Proyek ini adalah otomasi untuk evaluasi filter Kalman dan kuantisasi multibit (menggunakan Gray Code) terhadap data RSSI skenario 1 sampai 4. Skrip ini membaca file CSV sebagai *input*, melakukan kalkulasi praproses (seperti pencarian KGR, korelasi, dll), dan menyimpan hasilnya (beserta *arrays* bitstream) ke dalam beberapa file Excel *(sheet/workbook)* laporan yang dibedakan berdasarkan skenarionya. 

---

## 💻 Prasyarat
Pastikan Anda sudah menginstal **Python 3.8** atau versi lebih baru di perangkat Anda.

Library yang diperlukan untuk menjalankan program ini sudah tertulis di `requirements.txt`:
- numpy
- scipy
- pandas
- openpyxl

---

## 🚀 Cara Menjalankan Proyek

Berikut cara untuk mengonfigurasi *virtual environment* dan menjalankan proyek ini di sistem operasi **Windows** dan **Ubuntu (Linux)**.

### 🪟 Windows

1. **Buka Command Prompt atau PowerShell**, lalu arahkan (`cd`) ke folder proyek ini.
2. **Buat file Virtual Environment**:
   ```cmd
   python -m venv venv
   ```
3. **Aktifkan Virtual Environment**:
   ```cmd
   venv\Scripts\activate
   ```
   *(Tanda bahwa venv aktif adalah munculnya `(venv)` di sebelah kiri path command line Anda).*
4. **Instal seluruh library yang dibutuhkan**:
   ```cmd
   pip install -r requirements.txt
   ```
5. **Jalankan program utama**:
   ```cmd
   python main.py
   ```

### 🐧 Ubuntu (Linux)

1. **Buka Terminal**, lalu arahkan (`cd`) ke folder proyek ini.
2. **Buat file Virtual Environment** (Jika belum punya modul `venv`, install via `sudo apt install python3-venv`):
   ```bash
   python3 -m venv venv
   ```
3. **Aktifkan Virtual Environment**:
   ```bash
   source venv/bin/activate
   ```
   *(Tanda bahwa venv aktif adalah munculnya `(venv)` di sebelah kiri path terminal Anda).*
4. **Instal seluruh library yang dibutuhkan**:
   ```bash
   pip install -r requirements.txt
   ```
5. **Jalankan program utama**:
   ```bash
   python3 main.py
   ```

---

## 📁 Struktur Direktori Output

Setelah skrip sukses dijalankan, program akan menghasilkan sebuah folder `Output/` otomatis yang di dalamnya berisi:
* Skenario 1 sampai N (`skenario_1`, `skenario_2`, dst)
  * `data_excel_kalman/` - Berisi file array Kalman format excel.
  * `data_excel_kuantisasi/` - Berisi file bitstream hasil kuantisasi format excel.
  * `Laporan_Tabel_Kalman.xlsx` - Rekap tabel khusus evaluasi parameter Kalman.
  * `Laporan_Tabel_Kuantisasi.xlsx` - Rekap tabel khusus evaluasi parameter performansi (KDR, KGR, Komputasi).
# skg_automation
