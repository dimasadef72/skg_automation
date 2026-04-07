# Automation of Kalman Filter and Quantization

Berdasarkan dokumen `skenario mita.docx` dan umpan balik pengguna, kita menyederhanakan alur penyimpanan dan desain tabel agar **sama persis formatnya dengan output script bawaan**. 

## Analisis File Original
Setelah mengecek `KalmanFilter.py` dan `KuantisasiMultibit.py`, inilah 100% output asli mereka saat dijalankan MANUAL untuk 1 pasangan parameter:

**Dari KalmanFilter.py:**
- Menghasilkan array sinyal: `eve_alice_skenario1.csv`, `eve_bob_skenario1.csv`
- Menghitung MIN dan MAX? *(Di file Python saat ini, Min & Max belum diprogram, namun ada di permintaan dokumen Word).*
- Menyimpan **Tabel Benchmark 10 Iterasi** ke excel dengan kolom: `Iterasi` | `Waktu Alice` | `KGR Alice` | `Waktu Bob` | `KGR Bob`. Di baris terbawah ia menaruh nilai `Rata-rata` dan `Korelasi Pearson`.

**Dari KuantisasiMultibit.py:**
- Menghasilkan list bitstream: `alice_bitstream.csv`, `bob_bitstream.csv`
- Menyimpan **Tabel Benchmark 10 Iterasi** ke excel dengan kolom: `Percobaan` | `KDR (%)` | `KGR Alice` | `Time Alice` | `KGR Bob` | `Time Bob`

## Rancangan Struktur Rekomendasi Baru

Karena tabel asli menampung 10 percobaan per kondisi dan dipisah per pasangan (Alice-Bob terpisah dari EveAlice-EveBob), saya akan menyusun folder output automasi sedemikian rupa agar per parameternya menghasilkan tabel persis yang biasa Anda lihat. Semuanya diubah *extension*nya menjadi `.xlsx`.

```text
Output/
├── skenario_1/
│   ├── Q0.01_R0.5_bb1/                      # (Subfolder untuk 1 variasi parameter)
│   │   ├── kalman/
│   │   │   ├── excel/
│   │   │   │   ├── alice_kalman.xlsx        # File list nilai kalman
│   │   │   │   ├── bob_kalman.xlsx
│   │   │   │   ├── evealice_kalman.xlsx
│   │   │   │   └── evebob_kalman.xlsx
│   │   │   ├── kalman_alice_bob.xlsx        # Format tabel sama dg original
│   │   │   └── kalman_evealice_evebob.xlsx  # Format tabel sama dg original
│   │   └── kuantisasi/
│   │       ├── excel/
│   │       │   ├── alice_bitstream.xlsx     # File list string kuantisasi
│   │       │   ├── bob_bitstream.xlsx
│   │       │   ├── evealice_bitstream.xlsx
│   │       │   └── evebob_bitstream.xlsx
│   │       ├── kuantisasi_alice_bob.xlsx       # Format tabel sama dg original
│   │       └── kuantisasi_evealice_evebob.xlsx # Format tabel sama dg original
│   ├── Q0.01_R0.5_bb5/ ... (dan seterusnya hingga 8 variasi parameter)
│
├── skenario_2/ ...
├── skenario_3/ ...
└── skenario_4/ ...
```

### Tambahan Sesuai Skenario Asli (Min & Max)
Sesuai doc `skenario mita.docx`, dibutuhkan perhitungan nilai _Min dan Max_ untuk tiap sinyal Kalman. Karena di script asli belum ada fungsi *save* min max, saya akan menambahkan informasi Min & Max ini di baris terbawah pada file excel `kalman_alice_bob.xlsx` dan `kalman_evealice_evebob.xlsx`.

## Ringkasan Eksekusi `main.py`
Proses automasi secara program:
- Meloop skenario 1-4.
- Secara spesifik meloop iterasi konfigurasi (`Q, R, bb`).
- Membuka file sesuai lokasi asal, menjalankan evaluasi loop kecepatan/KGR 10x.
- Menyimpan data mentahnya ke dalam `.xlsx` list biasa.
- Menggenerasi file analisis Excel dengan 10 baris benchmark yang **mirip/identik dengan apa yang dihasilkan script original**, cukup dimodifikasi ke `.xlsx`.

## Open Questions

Apakah alur untuk *memisahkan* Excel evaluasi per variasi `(Q,R,bb)` & per pasangan komunikasi dan menjaga layout tabel *persis dengan aslinya* ini menyelesaikan kebingungan Anda? 

Bila Anda merasa struktur di atas (`Q0.01_R0.5_bb1`/ dst.. ) lebih rapi dan jelas, ketikkan "lanjut" dan implementasi kode pun segera saya buat tanpa merombak bentuk tabel asli Anda!
