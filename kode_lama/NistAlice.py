#!/usr/bin/env python3
"""
nist_equiv.py

Python equivalent of the C NIST test driver you provided.
- Input: CSV single-column file with one bit (0/1) per row (format A).
- Process: split into 128-bit blocks and run NIST-like tests per block:
  Approximate Entropy (m=3), Frequency, Block Frequency (M=3),
  Cumulative Sums (forward & reverse), Runs, Longest Runs of Ones.
- Output: prints p-values per key and writes ranking CSV (descending ApEn p-value)
  to files/sudahujinist_Alice.csv (1-based key indices).
"""

import os
import csv
import math
import numpy as np
import pandas as pd
from scipy import special
from scipy.stats import norm

# --- Config (edit paths if needed) ---
INPUT_CSV = "C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilHASH_eve/ALICE_HASH.csv"   # input single-column bit CSV
OUTPUT_RANK_CSV = "C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilNIST_eve/ALICE_NIST.csv"
BITS_PER_KEY = 128

# --- Helpers / NIST test implementations ---


def read_bits_single_col_csv(path):
    if not os.path.exists(path):
        raise FileNotFoundError(f"Input file not found: {path}")
    bits = []
    # read as CSV, first column only
    with open(path, newline="") as f:
        r = csv.reader(f)
        for row in r:
            if not row:
                continue
            cell = str(row[0]).strip()
            if cell == "":
                continue
            # if long bitstring, expand; else single bit
            if set(cell).issubset({"0", "1"}):
                if len(cell) == 1:
                    bits.append(int(cell))
                else:
                    bits.extend([int(ch) for ch in cell])
            else:
                # fallback numeric
                try:
                    v = int(float(cell))
                    bits.append(1 if v != 0 else 0)
                except Exception:
                    continue
    return bits


def approximate_entropy_test(epsilon, m=3):
    """Approximate Entropy test (NIST-like).
    Returns (p_value, apen_value) where apen_value = Phi(m) - Phi(m+1)
    """
    n = len(epsilon)
    # fallback if too short
    if n < 2 ** (m + 1):
        # reduce m until feasible (minimum m=1)
        while m > 1 and n < 2 ** (m + 1):
            m -= 1

    s = ''.join(str(b) for b in epsilon)

    def phi(mv):
        counts = {}
        pat_count = 2 ** mv
        # initialize counts
        for i in range(pat_count):
            p = bin(i)[2:].zfill(mv)
            counts[p] = 0
        # circular frequencies
        for i in range(n):
            pat = ''.join(s[(i + j) % n] for j in range(mv))
            counts[pat] += 1
        total = 0.0
        for c in counts.values():
            if c > 0:
                p = c / n
                total += p * math.log(p)
        return total

    phi_m = phi(m)
    phi_m1 = phi(m + 1)
    apen = phi_m - phi_m1
    chi_sq = 2.0 * n * (math.log(2) - apen)
    # degrees of freedom = 2^(m-1)
    df = 2 ** (m - 1)
    # p-value using upper incomplete gamma complement (gammaincc)
    p_value = special.gammaincc(df, chi_sq / 2.0)
    return p_value, apen


def frequency_test(epsilon):
    n = len(epsilon)
    s = sum(1 if bit == 1 else -1 for bit in epsilon)
    s_obs = abs(s) / math.sqrt(n)
    p_value = special.erfc(s_obs / math.sqrt(2))
    return p_value


def block_frequency_test(epsilon, M=3):
    n = len(epsilon)
    N = n // M
    if N == 0:
        return 0.0
    sum_v = 0.0
    for i in range(N):
        block = epsilon[i * M:(i + 1) * M]
        pi = sum(block) / M
        sum_v += (pi - 0.5) ** 2
    chi_sq = 4.0 * M * sum_v
    p_value = special.gammaincc(N / 2.0, chi_sq / 2.0)
    return p_value


def cumulative_sums_test(epsilon):
    """Return (p_forward, p_reverse) using NIST-like approach."""
    n = len(epsilon)
    x = [1 if b == 1 else -1 for b in epsilon]
    # forward
    S = 0
    sup = 0
    inf = 0
    for k in range(n):
        S += x[k]
        sup = max(sup, S)
        inf = min(inf, S)
    z = max(sup, -inf)
    if z == 0:
        # all zeros in partial sums -> p-value = 1.0 (no deviation)
        p_forward = 1.0
    else:
        # NIST formula using normal CDF (ndtr)
        sum1 = 0.0
        k_min = int(math.floor((-n / z + 1) / 4.0))
        k_max = int(math.floor((n / z - 1) / 4.0))
        for k in range(k_min, k_max + 1):
            sum1 += special.ndtr(((4 * k + 1) * z) / math.sqrt(n)) - special.ndtr(((4 * k - 1) * z) / math.sqrt(n))
        sum2 = 0.0
        k_min = int(math.floor((-n / z - 3) / 4.0))
        k_max = int(math.floor((n / z - 1) / 4.0))
        for k in range(k_min, k_max + 1):
            sum2 += special.ndtr(((4 * k + 3) * z) / math.sqrt(n)) - special.ndtr(((4 * k + 1) * z) / math.sqrt(n))
        p_forward = 1.0 - sum1 + sum2

    # reverse
    S = 0
    sup = 0
    inf = 0
    for k in range(n - 1, -1, -1):
        S += x[k]
        sup = max(sup, S)
        inf = min(inf, S)
    zrev = max(sup, -inf)
    if zrev == 0:
        p_reverse = 1.0
    else:
        sum1 = 0.0
        k_min = int(math.floor((-n / zrev + 1) / 4.0))
        k_max = int(math.floor((n / zrev - 1) / 4.0))
        for k in range(k_min, k_max + 1):
            sum1 += special.ndtr(((4 * k + 1) * zrev) / math.sqrt(n)) - special.ndtr(((4 * k - 1) * zrev) / math.sqrt(n))
        sum2 = 0.0
        k_min = int(math.floor((-n / zrev - 3) / 4.0))
        k_max = int(math.floor((n / zrev - 1) / 4.0))
        for k in range(k_min, k_max + 1):
            sum2 += special.ndtr(((4 * k + 3) * zrev) / math.sqrt(n)) - special.ndtr(((4 * k + 1) * zrev) / math.sqrt(n))
        p_reverse = 1.0 - sum1 + sum2

    return p_forward, p_reverse


def runs_test(epsilon):
    n = len(epsilon)
    S = sum(epsilon)
    pi = S / n
    # if condition fails, test not applicable (return 0.0 like C)
    if abs(pi - 0.5) > (2.0 / math.sqrt(n)):
        return 0.0
    # count runs
    V = 1
    for k in range(1, n):
        if epsilon[k] != epsilon[k - 1]:
            V += 1
    numerator = abs(V - 2.0 * n * pi * (1.0 - pi))
    denom = 2.0 * pi * (1.0 - pi) * math.sqrt(2.0 * n)
    erfc_arg = numerator / denom
    p_value = special.erfc(erfc_arg)
    return p_value


def longest_runs_test(epsilon):
    n = len(epsilon)
    if n < 128:
        return 0.0
    # parameters per NIST recommendations
    if n < 6272:
        K = 3
        M = 8
        V = [1, 2, 3, 4]
        pi = [0.21484375, 0.3671875, 0.23046875, 0.1875]
    elif n < 750000:
        K = 5
        M = 128
        V = [4, 5, 6, 7, 8, 9]
        pi = [0.1174035788, 0.242955959, 0.249363483, 0.17517706, 0.102701071, 0.112398847]
    else:
        K = 6
        M = 10000
        V = [10, 11, 12, 13, 14, 15, 16]
        pi = [0.0882, 0.2092, 0.2483, 0.1933, 0.1208, 0.0675, 0.0727]

    N = n // M
    nu = [0] * (K + 1)
    for i in range(N):
        block = epsilon[i * M:(i + 1) * M]
        longest_run = 0
        current = 0
        for bit in block:
            if bit == 1:
                current += 1
                if current > longest_run:
                    longest_run = current
            else:
                current = 0
        # categorize
        if longest_run <= V[0]:
            nu[0] += 1
        elif longest_run >= V[-1]:
            nu[K] += 1
        else:
            for j in range(len(V)):
                if longest_run == V[j]:
                    nu[j] += 1
                    break
    chi_sq = 0.0
    for i in range(K + 1):
        denom = N * pi[i]
        if denom > 0:
            chi_sq += ((nu[i] - N * pi[i]) ** 2) / denom
    p_value = special.gammaincc(K / 2.0, chi_sq / 2.0)
    return p_value


# --- Main driver that mimics your C program output ---
def test_universal_hash(input_file, output_file=None, bits_per_key=128):
    try:
        bits = read_bits_single_col_csv(input_file)
    except Exception as e:
        print(f"Error reading input: {e}")
        return None

    total_bits = len(bits)
    if total_bits < bits_per_key:
        print(f"Input too short ({total_bits} bits) for key length {bits_per_key}.")
        return None

    num_keys = total_bits // bits_per_key
    pvals_apen = []
    pvals_freq = []
    pvals_blockfreq = []
    pvals_cusumf = []
    pvals_cusumr = []
    pvals_runs = []
    pvals_longruns = []

    print(f"Created {num_keys} keys of {bits_per_key} bits each.")
    for k in range(num_keys):
        start = k * bits_per_key
        key_bits = bits[start:start + bits_per_key]
        print(f"=========== KEY {k + 1} ===========")
        # run tests
        p_apen, apen_val = approximate_entropy_test(key_bits, m=3)
        p_freq = frequency_test(key_bits)
        p_block = block_frequency_test(key_bits, M=3)
        p_cusumf, p_cusumr = cumulative_sums_test(key_bits)
        p_runs = runs_test(key_bits)
        p_long = longest_runs_test(key_bits)

        # print outputs (formatted similar to C)
        print(f"APPROXIMATE ENTROPY TEST [{k+1}]\t\t= {p_apen:.6f}")
        print(f"FREQUENCY TEST [{k+1}]\t\t\t= {p_freq:.6f}")
        print(f"BLOCK FREQUENCY TEST [{k+1}]\t\t= {p_block:.6f}")
        print(f"CUMULATIVE SUMS (FORWARD) TEST [{k+1}]\t= {p_cusumf:.6f}")
        print(f"CUMULATIVE SUMS (REVERSE) TEST [{k+1}]\t= {p_cusumr:.6f}")
        print(f"RUNS TEST [{k+1}]\t\t\t\t= {p_runs:.6f}")
        print(f"LONGEST RUNS OF ONES TEST [{k+1}]\t\t= {p_long:.6f}")
        print(f"APPROXIMATE ENTROPY VALUE [{k+1}]\t\t= {apen_val:.6f}")
        print("---------------------------------------------\n")

        pvals_apen.append(p_apen)
        pvals_freq.append(p_freq)
        pvals_blockfreq.append(p_block)
        pvals_cusumf.append(p_cusumf)
        pvals_cusumr.append(p_cusumr)
        pvals_runs.append(p_runs)
        pvals_longruns.append(p_long)

    # Ranking by approximate entropy (descending)
    indices = list(range(1, num_keys + 1))  # 1-based indexing for output like C
    # pair and sort by pvals_apen descending
    paired = list(zip(indices, pvals_apen))
    paired.sort(key=lambda x: x[1], reverse=True)
    ranked_keys = [idx for idx, pv in paired]

    print("\n\nRanking of keys by Approximate Entropy test (best first):")
    for rank, key_idx in enumerate(ranked_keys, start=1):
        print(f"Priority {rank}: Key {key_idx}")

    # write ranking file if requested
    if output_file:
        os.makedirs(os.path.dirname(output_file), exist_ok=True)
        with open(output_file, "w", newline='') as f:
            w = csv.writer(f)
            for idx in ranked_keys:
                w.writerow([idx])
        print(f"\n{output_file} file created")

    return ranked_keys


if __name__ == "__main__":
    # default run (same paths as your C code expectation)
    try:
        ranked = test_universal_hash(INPUT_CSV, OUTPUT_RANK_CSV, bits_per_key=BITS_PER_KEY)
    except Exception as ex:
        print("Fatal error:", ex)
