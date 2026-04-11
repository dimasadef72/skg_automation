#!/usr/bin/env python3
"""
Kuantisasi Multibit (Fixed / Non-Adaptive, Gray-Coded)
=====================================================
- Membaca dua file CSV (Alice & Bob)
- Mengkuantisasi data dengan jumlah bit tetap (uniform quantization)
- Menghasilkan bitstream Gray code untuk tiap data point
- Menghitung entropi, KGR, KDR
- Menyimpan hasil ke CSV dan Excel
"""

import os
import math
import time
import csv
import xlwt
import pandas as pd
import numpy as np
from typing import List

# ============================================================
# Utility functions
# ============================================================

def gray_code(n: int) -> List[str]:
    """Generate n-bit Gray code sequence."""
    if n <= 0:
        return ['0']
    codes = ['0', '1']
    for bits in range(2, n+1):
        mirror = codes[::-1]
        codes = ['0' + c for c in codes] + ['1' + c for c in mirror]
    return codes

def calculate_entropy(bitstream: str) -> float:
    """Shannon entropy of bitstream."""
    if not bitstream:
        return 0.0
    p0 = bitstream.count('0') / len(bitstream)
    p1 = bitstream.count('1') / len(bitstream)
    ent = 0.0
    if p0 > 0:
        ent -= p0 * math.log2(p0)
    if p1 > 0:
        ent -= p1 * math.log2(p1)
    return ent

def calculate_kdr(a: str, b: str) -> float:
    """Key Disagreement Rate (%)"""
    if not a or not b:
        return 0.0
    n = min(len(a), len(b))
    if n == 0:
        return 0.0
    diff = sum(1 for i in range(n) if a[i] != b[i])
    return (diff / n) * 100.0

# ============================================================
# Kuantisasi Multibit (Fixed / Non-Adaptive)
# ============================================================

def multibit_quantization(
    data: np.ndarray,
    num_bits: int = 3
):
    """
    Melakukan kuantisasi multibit tetap (uniform) dan menghasilkan bitstream Gray code.
    """
    t0 = time.perf_counter()
    data = np.asarray(data).astype(float)
    n = len(data)
    if n == 0:
        return "", 0, 0.0, 0.0, 0.0

    num_bits = int(max(1, min(8, num_bits)))
    levels = 2 ** num_bits

    data_min = np.min(data)
    data_max = np.max(data)
    data_range = data_max - data_min

    if data_range == 0:
        # Semua nilai sama → map ke satu level (0)
        indices = np.zeros(n, dtype=int)
    else:
        # Uniform quantization
        step = data_range / levels
        indices = np.floor((data - data_min) / step).astype(int)
        indices[indices >= levels] = levels - 1

    gray_map = gray_code(num_bits)
    bit_list = [gray_map[i] for i in indices]
    bitstream = "".join(bit_list)

    total_bits = len(bitstream)
    entropy = calculate_entropy(bitstream)

    t1 = time.perf_counter()
    elapsed = t1 - t0
    if elapsed == 0:
        elapsed = 1e-9
    kgr = total_bits / elapsed  # bit/s

    return bitstream, total_bits, entropy, kgr, elapsed

# ============================================================
# Benchmark
# ============================================================

def benchmark_kuantisasi(series: pd.Series, num_bits: int = 3, runs: int = 10):
    data = series.dropna().values.astype(float)
    bitstreams, times, entropies, kgrs, lengths = [], [], [], [], []

    for i in range(runs):
        bitstream, total_bits, entropy, kgr, t = multibit_quantization(data, num_bits=num_bits)
        print(f"[Run {i+1}] bits={total_bits}, entropy={entropy:.4f}, time={t:.6f}s, KGR={kgr:.2f}")
        bitstreams.append(bitstream)
        times.append(t)
        entropies.append(entropy)
        kgrs.append(kgr)
        lengths.append(total_bits)

    return bitstreams, times, entropies, kgrs, lengths

# ============================================================
# Output utilities
# ============================================================

def save_bitstream_to_csv(bitstream: str, path: str):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", newline='') as f:
        writer = csv.writer(f)
        writer.writerow(["bitstream"])
        writer.writerow([bitstream])
    print(f"✅ Bitstream disimpan ke: {path}")

def save_kdr_kgr_excel(
    bits_a, bits_b,
    kgrs_a, kgrs_b,
    times_a, times_b,
    outpath="Output/P2P/hasilkuantisasi_BBQbesar3/100eveanalisis_ringkas.xls"
):
    os.makedirs(os.path.dirname(outpath), exist_ok=True)
    book = xlwt.Workbook()
    sheet = book.add_sheet("KDR-KGR")

    headers = ["Percobaan", "KDR (%)", "KGR Alice", "Time Alice (s)", "KGR Bob", "Time Bob (s)"]
    for j, h in enumerate(headers):
        sheet.write(0, j, h)

    runs = min(len(bits_a), len(bits_b))
    for i in range(runs):
        kdr = calculate_kdr(bits_a[i], bits_b[i])
        sheet.write(i+1, 0, f"Run {i+1}")
        sheet.write(i+1, 1, round(kdr, 3))
        sheet.write(i+1, 2, round(kgrs_a[i], 2))
        sheet.write(i+1, 3, round(times_a[i], 6))
        sheet.write(i+1, 4, round(kgrs_b[i], 2))
        sheet.write(i+1, 5, round(times_b[i], 6))

    book.save(outpath)
    print(f"📊 File analisis disimpan ke: {outpath}")

# ============================================================
# Main CLI
# ============================================================

def read_first_column_csv(path: str) -> pd.Series:
    """Ambil kolom pertama dari CSV (header opsional)."""
    try:
        df = pd.read_csv(path, header=0)
        col = df.columns[0]
        series = pd.to_numeric(df[col], errors='coerce').dropna()
    except Exception:
        df = pd.read_csv(path, header=None)
        series = pd.to_numeric(df.iloc[:,0], errors='coerce').dropna()
    return series

def main():
    print("=== KUANTISASI MULTIBIT (NON-ADAPTIVE) ===\n")

    alice_path = input("Masukkan path file CSV untuk Alice: ").strip()
    bob_path   = input("Masukkan path file CSV untuk Bob  : ").strip()

    if not os.path.exists(alice_path) or not os.path.exists(bob_path):
        print("❌ File tidak ditemukan.")
        return

    df_a = read_first_column_csv(alice_path)
    df_b = read_first_column_csv(bob_path)

    num_bits = input("Masukkan jumlah bit per sample (1-8): ").strip()
    if not num_bits.isdigit():
        num_bits = 3
    else:
        num_bits = int(num_bits)
    runs = input("Masukkan jumlah percobaan (default 10): ").strip()
    runs = int(runs) if runs.isdigit() and int(runs) > 0 else 10

    out_dir = "Output/P2P/hasilkuantisasi_BBQbesar3"
    os.makedirs(out_dir, exist_ok=True)

    print(f"\n--- Alice (num_bits={num_bits}) ---")
    bits_a, times_a, ent_a, kgr_a, lens_a = benchmark_kuantisasi(df_a, num_bits=num_bits, runs=runs)
    save_bitstream_to_csv(bits_a[-1], os.path.join(out_dir, "100evealice_bitstream.csv"))

    print(f"\n--- Bob (num_bits={num_bits}) ---")
    bits_b, times_b, ent_b, kgr_b, lens_b = benchmark_kuantisasi(df_b, num_bits=num_bits, runs=runs)
    save_bitstream_to_csv(bits_b[-1], os.path.join(out_dir, "100evebob_bitstream.csv"))

    save_kdr_kgr_excel(bits_a, bits_b, kgr_a, kgr_b, times_a, times_b,
                      # os.path.join(out_dir, "analisis_ringkas.xls"))
    )
    print("\n✅ Proses selesai! Semua hasil ada di folder:")
    print(out_dir)

if __name__ == "__main__":
    main()
