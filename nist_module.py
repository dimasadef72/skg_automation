import time
import math

import numpy as np
from scipy import special


NIST_TEST_LABELS = [
    "Approximate Entropy",
    "Frequency",
    "Block Frequency",
    "Cumulative Sums (Forward)",
    "Cumulative Sums (Reverse)",
    "Runs",
    "Longest Runs of Ones",
]


def _hex_to_bits(hex_key):
    try:
        value = int(str(hex_key), 16)
    except ValueError:
        return []
    bits = bin(value)[2:].zfill(128)
    return [1 if bit == "1" else 0 for bit in bits]


def _pvalue_distribution(p_values):
    distribution = {
        "0.00-0.01": 0,
        "0.01-0.05": 0,
        "0.05-0.10": 0,
        "0.10-0.50": 0,
        "0.50-1.00": 0,
    }
    for p_value in p_values:
        if p_value < 0.01:
            distribution["0.00-0.01"] += 1
        elif p_value < 0.05:
            distribution["0.01-0.05"] += 1
        elif p_value < 0.10:
            distribution["0.05-0.10"] += 1
        elif p_value < 0.50:
            distribution["0.10-0.50"] += 1
        else:
            distribution["0.50-1.00"] += 1
    return distribution


def approximate_entropy_test(epsilon, m=3):
    n = len(epsilon)
    if n == 0:
        return 0.0, 0.0

    while m > 1 and n < 2 ** (m + 1):
        m -= 1

    bit_string = "".join(str(bit) for bit in epsilon)

    def phi(mv):
        counts = {}
        pattern_count = 2 ** mv
        for i in range(pattern_count):
            pattern = bin(i)[2:].zfill(mv)
            counts[pattern] = 0

        for i in range(n):
            pattern = "".join(bit_string[(i + j) % n] for j in range(mv))
            counts[pattern] += 1

        total = 0.0
        for count in counts.values():
            if count > 0:
                probability = count / n
                total += probability * math.log(probability)
        return total

    phi_m = phi(m)
    phi_m1 = phi(m + 1)
    apen = phi_m - phi_m1
    chi_sq = 2.0 * n * (math.log(2) - apen)
    degree_of_freedom = 2 ** (m - 1)
    p_value = special.gammaincc(degree_of_freedom, chi_sq / 2.0)
    return p_value, apen


def frequency_test(epsilon):
    n = len(epsilon)
    if n == 0:
        return 0.0
    s = sum(1 if bit == 1 else -1 for bit in epsilon)
    s_obs = abs(s) / math.sqrt(n)
    return special.erfc(s_obs / math.sqrt(2))


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
    return special.gammaincc(N / 2.0, chi_sq / 2.0)


def cumulative_sums_test(epsilon):
    n = len(epsilon)
    if n == 0:
        return 0.0, 0.0

    x = [1 if bit == 1 else -1 for bit in epsilon]

    def _directional_pvalue(sequence):
        total = 0
        sup = 0
        inf = 0
        for value in sequence:
            total += value
            sup = max(sup, total)
            inf = min(inf, total)

        z = max(sup, -inf)
        if z == 0:
            return 1.0

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

        return 1.0 - sum1 + sum2

    p_forward = _directional_pvalue(x)
    p_reverse = _directional_pvalue(list(reversed(x)))
    return p_forward, p_reverse


def runs_test(epsilon):
    n = len(epsilon)
    if n == 0:
        return 0.0

    s = sum(epsilon)
    pi = s / n
    if abs(pi - 0.5) > (2.0 / math.sqrt(n)):
        return 0.0

    V = 1
    for k in range(1, n):
        if epsilon[k] != epsilon[k - 1]:
            V += 1

    numerator = abs(V - 2.0 * n * pi * (1.0 - pi))
    denominator = 2.0 * pi * (1.0 - pi) * math.sqrt(2.0 * n)
    if denominator == 0:
        return 0.0

    return special.erfc(numerator / denominator)


def longest_runs_test(epsilon):
    n = len(epsilon)
    if n < 128:
        return 0.0

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
    if N == 0:
        return 0.0

    nu = [0] * (K + 1)
    for i in range(N):
        block = epsilon[i * M:(i + 1) * M]
        longest_run = 0
        current = 0
        for bit in block:
            if bit == 1:
                current += 1
                longest_run = max(longest_run, current)
            else:
                current = 0

        if longest_run <= V[0]:
            nu[0] += 1
        elif longest_run >= V[-1]:
            nu[K] += 1
        else:
            for j, threshold in enumerate(V):
                if longest_run == threshold:
                    nu[j] += 1
                    break

    chi_sq = 0.0
    for i in range(K + 1):
        denom = N * pi[i]
        if denom > 0:
            chi_sq += ((nu[i] - N * pi[i]) ** 2) / denom

    return special.gammaincc(K / 2.0, chi_sq / 2.0)


def process_nist(hash_keys, alpha=0.01):
    start = time.time()
    keys = list(hash_keys or [])

    test_results = {}
    for label in NIST_TEST_LABELS:
        test_results[label] = {
            "passed_count": 0,
            "avg_pvalue": 0.0,
            "pass_rate": 0.0,
            "distribution": {
                "0.00-0.01": 0,
                "0.01-0.05": 0,
                "0.05-0.10": 0,
                "0.10-0.50": 0,
                "0.50-1.00": 0,
            },
        }

    if not keys:
        end = time.time()
        return {
            "num_keys": 0,
            "passed_all_keys_count": 0,
            "tests": test_results,
            "time_nist": end - start,
        }

    p_values_by_test = {label: [] for label in NIST_TEST_LABELS}
    per_key = []
    passed_all_keys_count = 0

    for key in keys:
        bits = _hex_to_bits(key)
        test_pvalues = {
            "Approximate Entropy": approximate_entropy_test(bits, m=3)[0],
            "Frequency": frequency_test(bits),
            "Block Frequency": block_frequency_test(bits, M=3),
            "Cumulative Sums (Forward)": cumulative_sums_test(bits)[0],
            "Cumulative Sums (Reverse)": cumulative_sums_test(bits)[1],
            "Runs": runs_test(bits),
            "Longest Runs of Ones": longest_runs_test(bits),
        }

        if all(p_value >= alpha for p_value in test_pvalues.values()):
            passed_all_keys_count += 1

        for label, p_value in test_pvalues.items():
            p_values_by_test[label].append(p_value)

        per_key.append({"hex": key, "tests": test_pvalues})

    total_keys = len(keys)
    for label, p_values in p_values_by_test.items():
        passed_count = sum(1 for p_value in p_values if p_value >= alpha)
        test_results[label]["passed_count"] = passed_count
        test_results[label]["avg_pvalue"] = float(np.mean(p_values)) if p_values else 0.0
        test_results[label]["best_pvalue"] = float(np.max(p_values)) if p_values else 0.0
        test_results[label]["pass_rate"] = (passed_count / total_keys) * 100.0 if total_keys else 0.0
        test_results[label]["distribution"] = _pvalue_distribution(p_values)

    # find best key by Approximate Entropy (max p-value)
    best_key_by_apen = None
    if per_key:
        best_idx = max(range(len(per_key)), key=lambda i: per_key[i]["tests"]["Approximate Entropy"]) 
        best_key_by_apen = per_key[best_idx]["hex"]

    end = time.time()
    return {
        "num_keys": total_keys,
        "passed_all_keys_count": passed_all_keys_count,
        "tests": test_results,
        "per_key": per_key,
        "best_key_by_apen": best_key_by_apen,
        "time_nist": end - start,
    }
