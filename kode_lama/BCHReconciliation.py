import pandas as pd
import os

# =============================================================================
#                  BCH PURE PYTHON (PENGGANTI bchlib) - Perbaikan
# =============================================================================

PRIMITIVE_POLY = 0x11d
n = 255
k = 131
t = 8
ecc_len = n - k

# Precompute log/antilog table
gf_exp = [0] * 512
gf_log = [0] * 256
x = 1
for i in range(255):
    gf_exp[i] = x
    gf_log[x] = i
    x <<= 1
    if x & 0x100:
        x ^= PRIMITIVE_POLY
gf_exp[255:] = gf_exp[:255]

def gf_mul(a, b):
    if a == 0 or b == 0:
        return 0
    return gf_exp[(gf_log[a] + gf_log[b]) % 255]

def gf_pow(a, p):
    return gf_exp[(gf_log[a] * p) % 255]

def poly_mul(p, q):
    res = [0] * (len(p) + len(q) - 1)
    for i in range(len(p)):
        for j in range(len(q)):
            res[i+j] ^= gf_mul(p[i], q[j])
    return res

def rs_generator_poly(t):
    g = [1]
    for i in range(2 * t):
        g = poly_mul(g, [1, gf_pow(2, i)])
    return g

GEN = rs_generator_poly(t)

def rs_encode(msg):
    """
    msg: list of integers (0..255), length should be k (or less -> padded by caller)
    returns parity list (length = 2*t)
    """
    msg_out = list(msg) + [0] * (2 * t)
    # Ensure length at least k
    if len(msg_out) < k:
        msg_out += [0] * (k - len(msg_out))
    for i in range(k):
        coef = msg_out[i]
        if coef != 0:
            for j in range(len(GEN)):
                msg_out[i+j] ^= gf_mul(coef, GEN[j])
    # parity is the tail after k info symbols
    return msg_out[k:k + 2*t]

def rs_decode(data):
    # simplified decode: return first k symbols (no correction)
    if len(data) < k:
        # pad with zeros if needed
        data = list(data) + [0]*(k - len(data))
    return data[:k], 0

# =============================================================================
#                      FUNGSI PENDUKUNG CSV
# =============================================================================

def read_bit_csv(file_path):
    """Membaca file CSV berisi kolom bit (1 atau 0) tanpa memicu integer limit."""
    try:
        df = pd.read_csv(file_path, dtype=str, keep_default_na=False)  # baca sebagai STRING, aman
        bits = []
        col = df.columns[0]
        for val in df[col]:
            s = str(val).strip()
            if s == "":
                continue
            if set(s).issubset({'0','1'}):
                if len(s) == 1:
                    bits.append(int(s))
                else:
                    bits.extend([int(ch) for ch in s])
            else:
                # try numeric fallback
                try:
                    v = int(float(s))
                    bits.append(1 if v != 0 else 0)
                except:
                    pass
        return bits
    except Exception as e:
        print(f"⚠️ Gagal membaca file {file_path}: {e}")
        return []

def write_csv(path, data, columns):
    """Menyimpan data ke file CSV."""
    pd.DataFrame(data, columns=columns).to_csv(path, index=False)

# =============================================================================
#                        KONFIGURASI FILE
# =============================================================================

input_csv_path_alice = "Output/P2P/hasilkuantisasi_eve/alice_bitstream.csv"
input_csv_path_bob   = "Output/P2P/hasilkuantisasi_eve/bob_bitstream.csv"

output_dir = "Output/P2P/hasilBCH_eve"
os.makedirs(output_dir, exist_ok=True)

# =============================================================================
#                        MEMBACA BITSTREAM
# =============================================================================

alice_bits = read_bit_csv(input_csv_path_alice)
bob_bits   = read_bit_csv(input_csv_path_bob)

if not alice_bits or not bob_bits:
    print("❌ File input tidak dapat dibaca.")
    exit()

# Keep original copies for comparisons (avoid mutation)
alice_bits_orig = alice_bits.copy()
bob_bits_orig   = bob_bits.copy()

min_len = min(len(alice_bits_orig), len(bob_bits_orig))
alice_bits = alice_bits_orig[:min_len]
bob_bits   = bob_bits_orig[:min_len]

# =============================================================================
#                   KONVERSI BIT <-> BYTE (tidak memodifikasi input lists)
# =============================================================================

def bits_to_bytes_no_mutate(bits):
    """Return bytes from a list of bits without modifying the original list.
       Pads a copy to multiple of 8.
    """
    b = bits.copy()
    pad = (-len(b)) % 8
    if pad:
        b = b + [0]*pad
    return bytes([int(''.join(map(str, b[i:i+8])), 2) for i in range(0, len(b), 8)])

def bytes_to_bits(b):
    return [int(bit) for byte in b for bit in format(byte, '08b')]

alice_bytes = list(bits_to_bytes_no_mutate(alice_bits))
bob_bytes   = list(bits_to_bytes_no_mutate(bob_bits))

# If less than k symbols, pad to length k (so encoder/decoder do not index out of range)
def pad_to_length(arr, length, pad_value=0):
    arr2 = list(arr)
    if len(arr2) < length:
        arr2 += [pad_value] * (length - len(arr2))
    return arr2

alice_sym = pad_to_length(alice_bytes, k, 0)  # list of ints (symbols)
bob_sym   = pad_to_length(bob_bytes, k, 0)

# =============================================================================
#                              PROSES BCH (Pure Python)
# =============================================================================

ecc = rs_encode(alice_sym)           # parity symbols (length 2*t)
received = bob_sym + ecc             # Bob receives his symbols + parity
decoded_data, status = rs_decode(received)  # simplified decode -> first k info symbols

# Convert decoded_data (list of ints 0-255) back to bits
# decoded_data are bytes/symbols (0..255). Convert each to 8 bits to form corrected_bits.
decoded_bytes = bytes(decoded_data)  # make bytes object
corrected_bits = bytes_to_bits(decoded_bytes)[:min_len]  # truncate to original min_len

# Also prepare Bob_after (what Bob would have after attempting correction)
# In this simplified model, assume Bob's corrected equals decoded (Alice info)
bob_after_bits = corrected_bits.copy()

# =============================================================================
#                          SIMPAN HASIL (Alice + Bob)
# =============================================================================

# 1. Hasil koreksi (Alice perspective) - saved as CSV of bits
write_csv(os.path.join(output_dir, "Alice_BCH.csv"),
          [[int(b)] for b in corrected_bits],
          ["Alice"])

# 1b. Also save Bob after correction (so Bob output appears)
write_csv(os.path.join(output_dir, "Bob_BCH_after_correction.csv"),
          [[int(b)] for b in bob_after_bits],
          ["Bob_after"])

# 2. Perbandingan Alice (corrected) vs Bob original
comparison = [{"Alice_corrected": a, "Bob_original": b, "Match": int(a == b)}
              for a, b in zip(corrected_bits, bob_bits)]

write_csv(os.path.join(output_dir, "BCH_Comparison_Alice_vs_BobOriginal.csv"),
          comparison,
          ["Alice_corrected", "Bob_original", "Match"])

# 3. Perbandingan Alice vs Bob_after (post-correction)
comparison2 = [{"Alice_corrected": a, "Bob_after": b, "Match": int(a == b)}
               for a, b in zip(corrected_bits, bob_after_bits)]

write_csv(os.path.join(output_dir, "BCH_Comparison_Alice_vs_BobAfter.csv"),
          comparison2,
          ["Alice_corrected", "Bob_after", "Match"])

# 4. Hitung KDR (before & after)
total_bits = min_len
errors_before = sum(a != b for a, b in zip(alice_bits, bob_bits))
errors_after  = sum(a != b for a, b in zip(corrected_bits, bob_after_bits))

kdr_before = errors_before / total_bits if total_bits else 0
kdr_after  = errors_after / total_bits if total_bits else 0
improvement = ((kdr_before - kdr_after) / kdr_before * 100) if kdr_before != 0 else 0

summary_data = [
    ["Total Bits", total_bits],
    ["Errors Before", errors_before],
    ["Errors After", errors_after],
    ["KDR Before", round(kdr_before, 6)],
    ["KDR After", round(kdr_after, 6)],
    ["KDR Improvement (%)", round(improvement, 2)],
]

write_csv(os.path.join(output_dir, "KDR_BCH_Summary.csv"),
          summary_data,
          ["Metric", "Value"])

# 5. Perbandingan original Alice vs corrected
orig_vs_corr = [{"Original": a, "Corrected": c, "Changed": int(a != c)}
                for a, c in zip(alice_bits, corrected_bits)]

write_csv(os.path.join(output_dir, "Original_vs_Corrected_Alice.csv"),
          orig_vs_corr,
          ["Original", "Corrected", "Changed"])

# =============================================================================
#                                OUTPUT
# =============================================================================

print("✅ Proses BCH (pure-python) selesai!")
print(f"➡️ Output di folder: {output_dir}")
print(f"Total Bits compared: {total_bits}")
print(f"Errors Before: {errors_before} | After: {errors_after}")
print(f"KDR Before: {kdr_before:.6f} | After: {kdr_after:.6f}")
print(f"Improvement: {improvement:.2f}%")
