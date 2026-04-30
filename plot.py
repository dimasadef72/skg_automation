import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap
import matplotlib.patches as mpatches
import numpy as np
import os

def plot_bit_flow(csv_path):
    """
    Membaca file CSV yang berisi ['blok', 'bitstring']
    dan memvisualisasikannya dalam bentuk Grid Diagram Blok yang Berwarna.
    """
    if not os.path.exists(csv_path):
        print(f"[!] File tidak ditemukan: {csv_path}")
        return

    # 1. Baca data CSV
    try:
        df = pd.read_csv(csv_path)
    except Exception as e:
        print(f"[!] Gagal membaca CSV: {e}")
        return

    if 'bitstring' not in df.columns:
        print("[!] Format CSV salah: kolom 'bitstring' tidak ditemukan.")
        return

    # 2. Parsing Bitstring menjadi matriks numerik (0 dan 1)
    bit_matrix = []
    blok_labels = []
    
    for idx, row in df.iterrows():
        bit_str = str(row['bitstring']).strip()
        if not bit_str or bit_str.lower() == 'nan':
            continue
        
        # Ekstrak karakter '0' dan '1' menjadi list int
        bits = [int(b) for b in bit_str if b in ('0', '1')]
        
        if bits:
            bit_matrix.append(bits)
            blok_num = row['blok'] if 'blok' in df.columns else (idx + 1)
            blok_labels.append(blok_num)

    if not bit_matrix:
        print("[!] Tidak ada bit yang dapat diparsing dari data.")
        return

    # Padding jika panjang bit antar blok tidak sama
    max_len = max(len(b) for b in bit_matrix)
    bit_matrix_padded = [b + [np.nan] * (max_len - len(b)) for b in bit_matrix]
    arr2d = np.array(bit_matrix_padded)

    # 3. Setup Plot
    fig, axes = plt.subplots(2, 1, figsize=(14, 10))
    fig.patch.set_facecolor('#f4f4f4') # Latar belakang lebih lembut
    filename = os.path.basename(csv_path)

    # ---- Subplot 1: Grid Diagram Blok Berwarna ----
    ax1 = axes[0]
    
    # Custom warna: 0 = Biru (RoyalBlue), 1 = Kuning/Emas (Gold)
    cmap = ListedColormap(['#4169E1', '#FFD700'])
    cmap.set_bad(color='red') # Nan diwarnai merah (error padding)
    
    # Gunakan pcolormesh untuk efek kotak/blok yang lebih nyata dengan garis batas (edgecolors)
    # Pcolormesh butuh origin bawah-ke-atas, sehingga matriks dibalik ([::-1])
    # atau kita bisa gunakan imshow dengan minor gridlines tebal
    cax1 = ax1.imshow(arr2d, cmap=cmap, aspect='auto', interpolation='nearest')
    
    ax1.set_title(f"Diagram Blok Berwarna - Visualisasi per Blok\n{filename}", fontsize=15, weight='bold', color='#333333')
    ax1.set_ylabel("Nomor Blok", fontsize=12, weight='bold')
    ax1.set_xlabel("Posisi Bit dalam Blok", fontsize=12, weight='bold')
    
    # Mengatur label sumbu Y (Blok)
    step = max(1, len(blok_labels) // 20)
    ax1.set_yticks(range(0, len(blok_labels), step))
    ax1.set_yticklabels([blok_labels[i] for i in range(0, len(blok_labels), step)])

    # Garis pemisah antar blok (Grid warna putih yang membelah kotak-kotak)
    ax1.set_yticks(np.arange(-0.5, len(blok_labels), 1), minor=True)
    ax1.set_xticks(np.arange(-0.5, arr2d.shape[1], 1), minor=True)
    ax1.grid(which='minor', color='white', linestyle='-', linewidth=1.5)
    
    # Matikan tick lines agar terlihat lebih rapi
    ax1.tick_params(which='minor', bottom=False, left=False)
    
    # Legend Subplot 1
    patch0 = mpatches.Patch(color='#4169E1', label='Bit 0 (Biru)')
    patch1 = mpatches.Patch(color='#FFD700', label='Bit 1 (Emas)')
    ax1.legend(handles=[patch0, patch1], loc='upper right', bbox_to_anchor=(1.15, 1))

    # ---- Subplot 2: Barcode View Berwarna ----
    ax2 = axes[1]
    
    # Gabungkan semua blok menjadi 1 aliran 1D
    flat_bits = np.concatenate([np.array(b) for b in bit_matrix])
    
    # Batasi tampilan barcode agar plot tidak terlalu rapat (maksimal 500 bit pertama)
    max_display = min(500, len(flat_bits))
    display_bits = flat_bits[:max_display]
    
    x_pos = np.arange(len(display_bits))
    
    # Plot Bit 1 (Emas) dan Bit 0 (Biru)
    ax2.vlines(x_pos[display_bits == 1], ymin=0, ymax=1, color='#FFD700', linewidth=2.5)
    ax2.vlines(x_pos[display_bits == 0], ymin=0, ymax=1, color='#4169E1', linewidth=2.5)
    
    ax2.set_title(f"Visualisasi Barcode Aliran Bit ({max_display} Bit Pertama)", fontsize=15, weight='bold', color='#333333')
    ax2.set_xlabel("Posisi Bit Keseluruhan", fontsize=12, weight='bold')
    ax2.set_yticks([]) 
    ax2.set_xlim(-1, max_display)
    
    # Background untuk Subplot 2
    ax2.set_facecolor('#ffffff')

    plt.tight_layout()
    plt.show()


if __name__ == "__main__":
    # Path file CSV yang ingin divisualisasikan
    csv_file = r"D:\skg_automation\Output_aul\skenario_1\data_excel_bch\Q0.01_R0.5_BB100_alice_bch.csv"
    
    print(f"Mencoba memplot: {csv_file}")
    plot_bit_flow(csv_file)
