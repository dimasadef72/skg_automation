import time
import numpy as np
import math
from scipy.special import gammainc


def _hex_to_bits(hex_key):
    try:
        v = int(str(hex_key), 16)
    except ValueError:
        return []
    bits = bin(v)[2:].zfill(128)
    return [1 if b == '1' else 0 for b in bits]


# ──────────────────────────────────────────────────────────────
# Test 1: Frequency (Monobit) Test
# ──────────────────────────────────────────────────────────────
def _frequency_test(bits):
    n = len(bits)
    if n == 0:
        return 0.0
    s = sum(1 if b == 1 else -1 for b in bits)
    s_obs = abs(s) / math.sqrt(n)
    return math.erfc(s_obs / math.sqrt(2))


# ──────────────────────────────────────────────────────────────
# Test 2: Block Frequency Test
# ──────────────────────────────────────────────────────────────
def _block_frequency_test(bits, M=8):
    n = len(bits)
    if n < M:
        return 0.0
    N = n // M
    chi_sq = 0.0
    for i in range(N):
        block = bits[i * M:(i + 1) * M]
        pi_i = sum(block) / M
        chi_sq += (pi_i - 0.5) ** 2
    chi_sq *= 4 * M
    p_value = math.exp(-chi_sq / 2) if N == 1 else float(
        gammainc(N / 2.0, chi_sq / 2.0)
    )
    # use complementary incomplete gamma
    try:
        from scipy.special import gammaincc
        p_value = float(gammaincc(N / 2.0, chi_sq / 2.0))
    except Exception:
        pass
    return max(0.0, p_value)


# ──────────────────────────────────────────────────────────────
# Test 3: Cumulative Sums Test (forward)
# ──────────────────────────────────────────────────────────────
def _cumulative_sums_test(bits):
    n = len(bits)
    if n == 0:
        return 0.0
    x = [1 if b == 1 else -1 for b in bits]
    S = list(np.cumsum(x))
    z = max(abs(v) for v in S)
    if z == 0:
        return 1.0

    # Approximation formula from NIST SP 800-22
    def _sum_term(k_start, k_end, sign):
        total = 0.0
        for k in range(k_start, k_end + 1):
            from scipy.stats import norm
            total += (
                norm.cdf((4 * k + sign) * z / math.sqrt(n))
                - norm.cdf((4 * k - sign) * z / math.sqrt(n))
            )
        return total

    from scipy.stats import norm
    sum1 = 0.0
    for k in range(int(math.floor((-n / z + 1) / 4)), int(math.floor((n / z - 1) / 4)) + 1):
        sum1 += (
            norm.cdf((4 * k + 1) * z / math.sqrt(n))
            - norm.cdf((4 * k - 1) * z / math.sqrt(n))
        )
    sum2 = 0.0
    for k in range(int(math.floor((-n / z - 3) / 4)), int(math.floor((n / z - 1) / 4)) + 1):
        sum2 += (
            norm.cdf((4 * k + 3) * z / math.sqrt(n))
            - norm.cdf((4 * k + 1) * z / math.sqrt(n))
        )
    p_value = 1.0 - sum1 + sum2
    return max(0.0, min(1.0, float(p_value)))


# ──────────────────────────────────────────────────────────────
# Test 4: Runs Test
# ──────────────────────────────────────────────────────────────
def _runs_test(bits):
    n = len(bits)
    if n == 0:
        return 0.0
    ones = sum(bits)
    pi = ones / n
    if abs(pi - 0.5) >= (2.0 / math.sqrt(n)):
        return 0.0
    # Count runs
    V = 1 + sum(1 for i in range(1, n) if bits[i] != bits[i - 1])
    numer = abs(V - 2 * n * pi * (1 - pi))
    denom = 2 * math.sqrt(2 * n) * pi * (1 - pi)
    if denom == 0:
        return 0.0
    p_value = math.erfc(numer / denom)
    return max(0.0, p_value)


# ──────────────────────────────────────────────────────────────
# Test 5: Longest Run of Ones Test
# ──────────────────────────────────────────────────────────────
def _longest_run_test(bits):
    n = len(bits)
    if n < 128:
        return 0.0
    # Use M=8, K=3 for n in [128, 6271]
    M = 8
    # Theoretical probabilities for longest run in block of M=8
    pi = [0.2148, 0.3672, 0.2305, 0.1875]
    K = 3
    N = n // M
    # Count longest runs per block
    nu = [0] * (K + 1)
    for i in range(N):
        block = bits[i * M:(i + 1) * M]
        max_run = 0
        run = 0
        for b in block:
            if b == 1:
                run += 1
                max_run = max(max_run, run)
            else:
                run = 0
        v = min(max_run, K)
        # map: v<=1 -> 0, 2->1, 3->2, >=4->3
        if v <= 1:
            nu[0] += 1
        elif v == 2:
            nu[1] += 1
        elif v == 3:
            nu[2] += 1
        else:
            nu[3] += 1
    chi_sq = sum((nu[i] - N * pi[i]) ** 2 / (N * pi[i]) for i in range(K + 1))
    try:
        from scipy.special import gammaincc
        p_value = float(gammaincc(K / 2.0, chi_sq / 2.0))
    except Exception:
        p_value = 0.0
    return max(0.0, p_value)


# ──────────────────────────────────────────────────────────────
# Test 6: Approximate Entropy Test
# ──────────────────────────────────────────────────────────────
def _approx_entropy_test(bits, m=2):
    n = len(bits)
    if n < 10:
        return 0.0

    def _phi(m_val):
        counts = {}
        for i in range(n):
            # Circular template
            template = tuple(bits[(i + j) % n] for j in range(m_val))
            counts[template] = counts.get(template, 0) + 1
        c_vals = [v / n for v in counts.values()]
        return sum(c * math.log(c) for c in c_vals if c > 0)

    phi_m = _phi(m)
    phi_m1 = _phi(m + 1)
    ap_en = phi_m - phi_m1
    chi_sq = 2 * n * (math.log(2) - ap_en)
    try:
        from scipy.special import gammaincc
        p_value = float(gammaincc(2 ** (m - 1), chi_sq / 2.0))
    except Exception:
        p_value = 0.0
    return max(0.0, p_value)


# ──────────────────────────────────────────────────────────────
# Main process_nist: runs all 6 tests, picks highest p-value
# ──────────────────────────────────────────────────────────────
def process_nist(hash_keys, alpha=0.01):
    """
    Menjalankan 6 uji NIST SP 800-22 untuk setiap kunci:
      1. Frequency Test
      2. Block Frequency Test
      3. Cumulative Sums Test
      4. Runs Test
      5. Longest Run of Ones Test
      6. Approximate Entropy Test

    Untuk setiap kunci, diambil nilai p-value tertinggi dari semua 6 uji.
    Fungsi mengembalikan:
      - passed_keys_count  : jumlah kunci yang lulus (p >= alpha)
      - avg_pvalue         : rata-rata p-value tertinggi
      - pass_rate          : persentase kunci lulus
      - distribution       : distribusi p-value tertinggi
      - time_nist          : waktu komputasi (detik)
      - test_pvalues_best  : dict {test_name: avg_best_pvalue} untuk tiap test
    """
    start = time.time()
    passed_keys_count = 0
    avg_pvalue = 0.0

    _empty_dist = {
        "0.00-0.01": 0,
        "0.01-0.05": 0,
        "0.05-0.10": 0,
        "0.10-0.50": 0,
        "0.50-1.00": 0,
    }
    _empty_tests = {
        "Frequency": 0.0,
        "Block Frequency": 0.0,
        "Cumulative Sums": 0.0,
        "Runs": 0.0,
        "Longest Run": 0.0,
        "Approx Entropy": 0.0,
    }

    if not hash_keys:
        return passed_keys_count, avg_pvalue, 0.0, _empty_dist, 0.0, _empty_tests

    TEST_NAMES = [
        ("Frequency",      _frequency_test),
        ("Block Frequency", _block_frequency_test),
        ("Cumulative Sums", _cumulative_sums_test),
        ("Runs",           _runs_test),
        ("Longest Run",    _longest_run_test),
        ("Approx Entropy", _approx_entropy_test),
    ]

    best_pvalues = []           # p-value tertinggi per kunci
    per_test_sums = {name: 0.0 for name, _ in TEST_NAMES}

    for key in hash_keys:
        bits = _hex_to_bits(key)
        pvals_this_key = {}
        for t_name, t_func in TEST_NAMES:
            try:
                pv = t_func(bits)
            except Exception:
                pv = 0.0
            pvals_this_key[t_name] = pv
            per_test_sums[t_name] += pv

        best_p = max(pvals_this_key.values())
        best_pvalues.append(best_p)

    n_keys = len(hash_keys)
    if best_pvalues:
        avg_pvalue = float(np.mean(best_pvalues))
        passed_keys_count = sum(1 for p in best_pvalues if p >= alpha)

    pass_rate = (passed_keys_count / n_keys) * 100.0

    distribution = {
        "0.00-0.01": 0,
        "0.01-0.05": 0,
        "0.05-0.10": 0,
        "0.10-0.50": 0,
        "0.50-1.00": 0,
    }
    for p in best_pvalues:
        if p < 0.01:
            distribution["0.00-0.01"] += 1
        elif p < 0.05:
            distribution["0.01-0.05"] += 1
        elif p < 0.10:
            distribution["0.05-0.10"] += 1
        elif p < 0.50:
            distribution["0.10-0.50"] += 1
        else:
            distribution["0.50-1.00"] += 1

    # Rata-rata per uji
    test_avg_pvalues = {name: (per_test_sums[name] / n_keys) for name in per_test_sums}

    end = time.time()
    time_nist = end - start

    return passed_keys_count, avg_pvalue, pass_rate, distribution, time_nist, test_avg_pvalues
