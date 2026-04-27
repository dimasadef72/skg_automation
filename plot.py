import pandas as pd
import matplotlib.pyplot as plt

# ------------------------------------------------------------
# 1. Membaca file CSV
# ------------------------------------------------------------
file_path = 'D:/SKG_TEST/skenario2.csv'
df = pd.read_csv(file_path)

# ------------------------------------------------------------
# 2. Mengambil kolom data RSSI
# ------------------------------------------------------------
rssi_alice = df['BOB']
rssi_bob = df['ALICE']
rssi_alice_eve = df['EVE_BOB']
rssi_bob_eve = df['EVE_ALICE']

# ------------------------------------------------------------
# 3. Plot grafik RSSI
# ------------------------------------------------------------
plt.figure(figsize=(14, 7))

plt.plot(rssi_alice, marker='o', linewidth=1.5, label="Alice → Bob (RSSI)")
plt.plot(rssi_bob, marker='o', linewidth=1.5, label="Bob → Alice (RSSI)")
plt.plot_rssi_alice_eve = plt.plot(rssi_alice_eve, marker='o', linewidth=1.5, label="Alice → Eve (RSSI)")
plt.plot(rssi_bob_eve, marker='o', linewidth=1.5, label="Bob → Eve (RSSI)")

# Judul plot
plt.title("Skenario 2 – Perbandingan Nilai RSSI Antar Perangkat", fontsize=14, weight='bold')

# Label sumbu
plt.xlabel("Index Sampel", fontsize=12)
plt.ylabel("RSSI (dBm)", fontsize=12)

# Atur batas sumbu vertikal (misalnya dari -100 sampai -40)
plt.ylim(-100, -20)

# Grid lebih tegas
plt.grid(True, linestyle='--', linewidth=0.6, alpha=0.7)

# Legend
plt.legend(fontsize=10)

plt.tight_layout()
plt.show()
