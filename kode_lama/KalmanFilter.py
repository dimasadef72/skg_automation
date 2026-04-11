import numpy as np
import time
import csv
import os
from openpyxl import Workbook
from scipy.stats import pearsonr

# === Parameter Kalman ===
a = 1
h = 1
R = 0.5 #dibawah 1
Q = 0.01
xaposteriori_0 = -5
paposteriori_0 = 1
bb = 50 #jumlah measurement per sample
interval = 0.11

# === Baca data dari file CSV terpisah ===
def read_rssi_csv(path):
    data = []
    with open(path, 'r') as f:
        reader = csv.reader(f)
        for row in reader:
            try:
                data.append(int(row[0]))  # baca kolom pertama
            except:
                continue  # lewati baris kosong atau header
    return data

rss_alice = read_rssi_csv('skenario1_mita_alice.csv')
rss_bob   = read_rssi_csv('skenario1_mita_bob.csv')

# === Persiapan reshaping ===
total_data = min(len(rss_alice), len(rss_bob))  # pastikan panjang sama
aa = total_data // bb
rss_alice = rss_alice[:aa * bb]
rss_bob = rss_bob[:aa * bb]
alice = np.array(rss_alice).reshape(aa, bb).T
bob = np.array(rss_bob).reshape(aa, bb).T

# === Fungsi Kalman Filter ===
def kalman_filter(signal):
    xaposteriori = []
    paposteriori = []
    row1=[]; row2=[]; row3=[]; row4=[]; row5=[]; row6=[]

    for m in range(aa):
        row1.append(a * xaposteriori_0)
        row2.append(signal[0][m] - h * row1[m])
        row3.append(a*a * paposteriori_0 + Q)
        gain = row3[m] / (row3[m] + R)
        row4.append(gain)
        row5.append(row3[m] * (1 - gain))
        row6.append(row1[m] + gain * row2[m])
    xaposteriori.append(row6)
    paposteriori.append(row5)

    for j in range(1, bb):
        r1=[]; r2=[]; r3=[]; r4=[]; r5=[]; r6=[]
        for n in range(aa):
            r1.append(xaposteriori[j-1][n])
            r2.append(signal[j][n] - h * r1[n])
            r3.append(a*a * paposteriori[j-1][n] + Q)
            gain = r3[n] / (r3[n] + R)
            r4.append(gain)
            r5.append(r3[n] * (1 - gain))
            r6.append(r1[n] + gain * r2[n])
        xaposteriori.append(r6)
        paposteriori.append(r5)
    return xaposteriori

# === Buat direktori output ===
os.makedirs('Output/P2P/hasilpraproses_BB33', exist_ok=True)

# === Simpan hasil Kalman Alice ===
hasil_alice = np.array(kalman_filter(alice)).T.reshape(-1, 1)
with open('Output/P2P/hasilpraproses_BB33/100evealice_process_skenario4.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['Alice_praproses'])
    for val in hasil_alice:
        writer.writerow([int(val.item())])

# === Simpan hasil Kalman Bob ===
hasil_bob = np.array(kalman_filter(bob)).T.reshape(-1, 1)
with open('Output/P2P/hasilpraproses_BB33/100evebob_praprocess_skenario4.csv', 'w', newline='') as f:
    writer = csv.writer(f)
    writer.writerow(['Bob_praproses'])
    for val in hasil_bob:
        writer.writerow([int(val.item())])

# === Korelasi Pearson ===
korelasi, _ = pearsonr(hasil_alice.flatten(), hasil_bob.flatten())

# === Benchmark & KGR ===
def benchmark(signal, label):
    times = []
    kgrs = []

    for i in range(10):
        start = time.perf_counter()
        result = kalman_filter(signal)
        end = time.perf_counter()
        elapsed = end - start
        gain_count = aa * bb
        kgr = gain_count * 32 / elapsed
        print(f"{label} percobaan ke-{i+1}: {elapsed:.6f} detik | KGR: {kgr:.2f} bit/s")
        times.append(elapsed)
        kgrs.append(kgr)

    avg_time = sum(times) / len(times)
    avg_kgr = sum(kgrs) / len(kgrs)
    print(f"Rata-rata waktu Kalman {label}: {avg_time:.6f} detik")
    print(f"Rata-rata KGR Kalman {label}: {avg_kgr:.2f} bit/s\n")
    return times, kgrs, avg_time, avg_kgr

# === Jalankan benchmark ===
print("===== BENCHMARK ALICE =====")
times_alice, kgr_alice, avg_time_alice, avg_kgr_alice = benchmark(alice, 'Alice')

print("===== BENCHMARK BOB =====")
times_bob, kgr_bob, avg_time_bob, avg_kgr_bob = benchmark(bob, 'Bob')

# === Simpan hasil analisis ke Excel ===
wb = Workbook()
ws = wb.active
ws.title = "100eveAnalisis Kalman 4"

ws.append(["Iterasi", "Waktu Alice (s)", "KGR Alice (bit/s)", "Waktu Bob (s)", "KGR Bob (bit/s)"])

for i in range(10):
    ws.append([i + 1, times_alice[i], kgr_alice[i], times_bob[i], kgr_bob[i]])

for row in ws.iter_rows(min_row=2, max_row=11, min_col=2, max_col=5):
    for cell in row:
        cell.number_format = '0.000000'

ws.append(["RATA-RATA", avg_time_alice, avg_kgr_alice, avg_time_bob, avg_kgr_bob])
ws.append([])
ws.append(["Korelasi Pearson", korelasi])

excel_path = 'Output/P2P/hasilpraproses_BB33/100eveanalisis_kalman4.xlsx'
wb.save(excel_path)
print(f"Hasil analisis disimpan ke {excel_path}")
