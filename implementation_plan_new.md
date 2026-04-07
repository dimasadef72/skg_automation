# Automation of Kalman Filter and Quantization

Berdasarkan gambar tabel dari dokumen/tesis yang Anda lampirkan pelajari, saya akhirnya memahami format yang Anda inginkan. Script Python ini (`main.py`) tidak lagi akan menghasilkan baris per baris data mentah iterasi, melainkan **akan secara otomatis menghasilkan/menggambar tabel persis seperti screenshot dokumen Anda (dengan cell yang di-merge) langsung ke Excel!**

## Visualisasi Output Tabel (Sesuai Screenshot)

### 1. Tabel Kalman (Praproses)
Output `tabel_kalman_skenario_X.xlsx` akan terbuat menggunakan `openpyxl`. Setiap variasi (contoh `Q=0.01, R=0.5, bb=1`) akan menghasilkan 1 blok tabel seperti ini:

| Parameter | Sebelum Praproses (Alice) | Sebelum Praproses (Bob) | Sebelum (Eve-Alice) | Sebelum (Eve-Bob) | Setelah Praproses (Alice) | Setelah Praproses (Bob) | Setelah (Eve-Alice) | Setelah (Eve-Bob) |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Maksimum (dBm)**| -29 | -27 | -33 | -28 | -21 | -19 | -23 | -20 |
| **Minimum (dBm)** | -83 | -83 | -89 | -25 | -57 | -57 | -61 | -25 |
| **Koefisien Korelasi** | colspan=2 (Korelasi A&B Sebelum) | colspan=2 (Korelasi E-A & E-B Sblm) | colspan=2 (Korelasi A&B Setelah) | colspan=2 (Korelasi E-A & E-B Setelah) |
| **Waktu Komputasi (s)** | colspan=4 (Kosong) | Waktu A | Waktu B | Waktu E-A | Waktu E-B |
| **KGR (bit/s)** | colspan=4 (Kosong) | KGR A | KGR B | KGR E-A | KGR E-B |

### 2. Tabel Kuantisasi
Output `tabel_kuantisasi_skenario_X.xlsx` akan terbangun rapi dengan header _Parameter Performansi_ seperti ini:

| Parameter Performansi | Alice | Bob | Eve-Alice | Eve-Bob |
| :--- | :--- | :--- | :--- | :--- |
| **KDR (%)** | colspan=2 (KDR A&B) | colspan=2 (KDR E-A & E-B) |
| **KGR (bit/s)** | KGR A | KGR B | KGR E-A | KGR E-B |
| **Waktu komputasi (s)** | Waktu A | Waktu B | Waktu E-A | Waktu E-B |

*(Catatan: Rata-rata waktu komputasi & KGR akan tetap dihitung secara statis dari rata-rata 10x loop di latar belakang agar hasilnya seakurat mungkin, tetapi yang ditampilkan adalah tabel bersih ini)*.

## Rancangan Struktur Automasi

Karena Anda akan perlu laporan untuk dokumen Anda secara utuh, saya akan merombak strukturnya menjadi sangat rapi untuk Microsoft Word Anda:

```text
Output/
├── skenario_1/
│   ├── data_excel/
│   │   ├── (Berisi output data mentah array/bitstream `.xlsx` untuk seluruh perhitungan bila Anda butuh melihat sinyalnya)
│   ├── Laporan_Tabel_Kalman.xlsx       (Berisi 8 block tabel Kalman yang sudah tergambar merged/diformat warnanya)
│   └── Laporan_Tabel_Kuantisasi.xlsx   (Berisi 8 block tabel Kuantisasi yang sudah tergambar cell formatnya)
│
├── skenario_2/...
├── skenario_3/...
└── skenario_4/...
```

Kelebihan utama dari struktur ini:
Anda **tidak perlu mengecilkan/menggabung (merge) data secara manual lagi**. Buka Excel-nya, copy tabelnya, paste ke Word—selesai. 

## Flow Program `main.py`
Proses yang terjadi di script Python:
- Ekstraksi matrix sinyal asli (Sebelum Praproses). Fungsi tambahan akan diaktifkan untuk mencari Min, Max, dan korelasi data asli.
- Pembuatan Kalman Filter dari *logic* `KalmanFilter.py` lama dan perhitungan Waktu, KGR, Min, Max (Setelah Praproses).
- Pembuatan hasil bitstream menggunakan *logic* `KuantisasiMultibit.py`, perhitungan parameter, disisipkan ke dalam format tabel openpyxl ter-merge.

## Open Question

Terima kasih sekali atas foto ini, sekarang arahan Anda sudah 100% sangat jelas. Skrip ini akan **menggambar cel excel** yang sudah berformat persis seperti gambar Anda. 

Silakan cek rancangan tabel dan susunan foldernya di atas! Bila kita sudah satu gelombang dan ini yang Anda mau, bilang "**setuju bikin kodenya**", dan saya akan buat script automasinya.
