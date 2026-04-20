import time
import numpy as np
import math


def _hex_to_bits(hex_key):
    try:
        v = int(str(hex_key), 16)
    except ValueError:
        return []
    bits = bin(v)[2:].zfill(128)
    return [1 if b == '1' else 0 for b in bits]


def _frequency_pvalue(bits):
    n = len(bits)
    if n == 0:
        return 0.0
    s = sum(1 if bit == 1 else -1 for bit in bits)
    s_obs = abs(s) / math.sqrt(n)
    return math.erfc(s_obs / math.sqrt(2))


def process_nist(hash_keys, alpha=0.01):
    start = time.time()
    passed_keys_count = 0
    avg_pvalue = 0.0

    if not hash_keys:
        distribution = {
            "0.00-0.01": 0,
            "0.01-0.05": 0,
            "0.05-0.10": 0,
            "0.10-0.50": 0,
            "0.50-1.00": 0,
        }
        return passed_keys_count, avg_pvalue, 0.0, distribution, 0.0

    p_values = []
    for key in hash_keys:
        bits = _hex_to_bits(key)
        p_values.append(_frequency_pvalue(bits))

    if p_values:
        avg_pvalue = float(np.mean(p_values))
        passed_keys_count = sum(1 for p in p_values if p >= alpha)

    pass_rate = (passed_keys_count / len(hash_keys)) * 100.0 if hash_keys else 0.0

    distribution = {
        "0.00-0.01": 0,
        "0.01-0.05": 0,
        "0.05-0.10": 0,
        "0.10-0.50": 0,
        "0.50-1.00": 0,
    }
    for p in p_values:
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

    end = time.time()
    time_nist = end - start

    return passed_keys_count, avg_pvalue, pass_rate, distribution, time_nist
