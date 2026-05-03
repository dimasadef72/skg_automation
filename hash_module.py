import time
import numpy as np
import math
import hashlib


def process_hash(alice_bits, bob_bits, hashtable_path='Hashtable128.csv'):
    """
        Proses Privacy Amplification dengan 3 sub-tahap terpisah:
      1. Universal Hash  -> (keys_alice, keys_bob, hex_alice, hex_bob, time_universal)
      2. SHA-1           -> (sha_keys_alice, sha_keys_bob, time_sha)
      3. AES-128 Matching -> (aes_keys, time_aes)

    Return:
            hex_alice, hex_bob,
            sha_keys_alice, sha_keys_bob,
            aes_keys,
            time_universal, time_sha, time_aes,
      metrics
    """

    # ──────────────────────────────────────────────────────────────
    # Baca Hashtable
    # ──────────────────────────────────────────────────────────────
    try:
        Hashtab = []
        with open(hashtable_path, newline='') as f:
            import csv
            reader = csv.reader(f)
            for row in reader:
                Hashtab.append(row)
        Hashtable = np.array(Hashtab, dtype=int)
    except FileNotFoundError:
        print(f"Warning: {hashtable_path} tidak ditemukan. Menggunakan dummy hashtable.")
        Hashtable = np.ones((128, 128), dtype=int)

    # ──────────────────────────────────────────────────────────────
    # Tahap 1: Universal Hash
    # ──────────────────────────────────────────────────────────────
    t_univ_start = time.time()

    def univ_hash(bits):
        ukuranhash = 128
        ln = len(bits) - (len(bits) % 128)
        bits_cut = bits[:ln]
        jumlahkey = math.floor(len(bits_cut) / 128)

        keys = []
        for i in range(jumlahkey):
            aaa = bits_cut[(i * ukuranhash): (ukuranhash * (i + 1))]
            mat1 = []
            for x in range(len(Hashtable)):
                total = 0
                for y in range(len(aaa)):
                    total += Hashtable[x][y] * aaa[y]
                mat1.append(int(total % 2))
            keys.append(mat1)
        return keys

    keys_alice = univ_hash(alice_bits)
    keys_bob   = univ_hash(bob_bits)

    # Konversi ke hex (diperlukan untuk NIST setelah universal hash)
    def to_hex(keys_list):
        hex_list = []
        for k in keys_list:
            keybit = "".join(str(e) for e in k)
            keyint = int(keybit, 2)
            hex_list.append("%032x" % keyint)
        return hex_list

    hex_alice = to_hex(keys_alice)
    hex_bob   = to_hex(keys_bob)

    t_univ_end = time.time()
    time_universal = t_univ_end - t_univ_start

    # ──────────────────────────────────────────────────────────────
    # Tahap 2: SHA-1
    # ──────────────────────────────────────────────────────────────
    t_sha_start = time.time()

    sha_keys_alice = []
    for k in keys_alice:
        hash1 = k * 1
        data1 = "".join(str(e) for e in hash1)
        someText1 = data1.encode("ascii")
        sha_keys_alice.append(hashlib.sha1(someText1).hexdigest())

    sha_keys_bob = []
    for k in keys_bob:
        hash1 = k * 1
        data1 = "".join(str(e) for e in hash1)
        someText1 = data1.encode("ascii")
        sha_keys_bob.append(hashlib.sha1(someText1).hexdigest())

    t_sha_end = time.time()
    time_sha = t_sha_end - t_sha_start

    # ──────────────────────────────────────────────────────────────
    # Tahap 3: AES-128 Key Matching
    # ──────────────────────────────────────────────────────────────
    t_aes_start = time.time()

    aes_keys = []
    for idx in range(min(len(sha_keys_alice), len(sha_keys_bob))):
        if sha_keys_alice[idx] == sha_keys_bob[idx]:
            # Menggunakan hasil Universal Hash (128-bit / 32 karakter hex) sebagai kunci AES-128
            aes_keys.append(hex_alice[idx])

    t_aes_end = time.time()
    time_aes = t_aes_end - t_aes_start

    # ──────────────────────────────────────────────────────────────
    # Metrics
    # ──────────────────────────────────────────────────────────────
    metrics = {
        "input_bits_alice":     len(alice_bits),
        "input_bits_bob":       len(bob_bits),
        "keys_count_alice":     len(keys_alice),
        "keys_count_bob":       len(keys_bob),
        "total_key_bits_alice": len(keys_alice) * 128,
        "total_key_bits_bob":   len(keys_bob) * 128,
        "matched_key_count":    len(aes_keys),
        "matched_key_bits":     len(aes_keys) * 128,
        "time_universal":       time_universal,
        "time_sha":             time_sha,
        "time_aes":             time_aes,
        # backward-compat key (total)
        "time_hash":            time_universal + time_sha + time_aes,
    }

    return (
        hex_alice, hex_bob,
        sha_keys_alice, sha_keys_bob,
        aes_keys,
        time_universal, time_sha, time_aes,
        metrics
    )
