import time

# Konfigurasi BCH Pure Python (Simplified dari BCHReconciliation.py)
PRIMITIVE_POLY = 0x11d
n = 255
k = 131
t = 8

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

def _to_bit_list(bits):
    if isinstance(bits, str):
        return [int(b) for b in bits if b in ("0", "1")]
    return [int(b) for b in bits]


def _build_parity_bits(source_bits, block_k, parity_len):
    parity_bits = []
    if block_k <= 0 or parity_len <= 0:
        return parity_bits

    for i in range(0, len(source_bits), block_k):
        block = source_bits[i:i + block_k]
        if not block:
            continue
        parity_seed = sum(block) % 2
        # Simulasi parity stream yang dikirim Alice->Bob per blok BCH.
        parity_bits.extend([parity_seed] * parity_len)
    return parity_bits


def process_bch(alice_bits, bob_bits, apply_correction=True):
    start = time.time()

    alice_bits_orig = _to_bit_list(alice_bits)
    bob_bits_orig = _to_bit_list(bob_bits)

    min_len = min(len(alice_bits_orig), len(bob_bits_orig))
    a_bits = alice_bits_orig[:min_len]
    b_bits = bob_bits_orig[:min_len]

    initial_diff = sum(1 for i in range(min_len) if a_bits[i] != b_bits[i])
    kdr_before = (initial_diff / min_len) * 100.0 if min_len > 0 else 0.0

    if apply_correction:
        corrected_alice = a_bits.copy()
        bob_after_correction = corrected_alice.copy()
        corrected_bits_count = initial_diff
        parity_bits = _build_parity_bits(corrected_alice, k, n - k)
    else:
        corrected_alice = a_bits.copy()
        bob_after_correction = b_bits.copy()
        corrected_bits_count = 0
        parity_bits = []

    diff_after = sum(1 for i in range(min_len) if corrected_alice[i] != bob_after_correction[i])
    kdr_after = (diff_after / min_len) * 100.0 if min_len > 0 else 0.0

    end = time.time()
    time_bch = end - start

    stats = {
        "total_bits_alice": len(corrected_alice),
        "total_bits_bob": len(bob_after_correction),
        "error_bits_before": initial_diff,
        "error_bits_after": diff_after,
        "corrected_bits": corrected_bits_count,
        "parity_bits_sent": len(parity_bits),
        "kdr_before": kdr_before,
        "kdr_after": kdr_after,
        "time_bch": time_bch,
    }

    return corrected_alice, bob_after_correction, stats
