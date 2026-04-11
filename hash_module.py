import time
import numpy as np
import math
import hashlib

def process_hash(alice_bits, bob_bits, hashtable_path='Hashtable128.csv'):
    start = time.time()
    
    # 1. Baca Hashtable
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

    # Bikin Hash Alice
    keys_alice = univ_hash(alice_bits)
    # Bikin Hash Bob
    keys_bob = univ_hash(bob_bits)
    
    # Bikin SHA & AES 
    sha_keys_alice = []
    for k in keys_alice:
        # Kode lama Anda: `hash1 = [k] * 128` ini salah karena mereplikasi list k 128 kali,
        # seharusnya mereplikasi elemen k, karena k adalah list [0,1,0...]
        # Di kode asli: `hash1 = [key1[0][valid_indeks[j]]] * 128` (mengemas seluruh blok 128 bit)
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

    # === BIKIN KUNCI HEX UNTUK SEMUA (ALICE & BOB) ===
    # Dibikin dulu biar bisa dilihat biarpun beda/salah
    def to_hex(keys_list):
        hex_list = []
        for k in keys_list:
            keybit = "".join(str(e) for e in k)
            keyint = int(keybit, 2)
            hex_list.append("%032x" % keyint)
        return hex_list
        
    hex_alice = to_hex(keys_alice)
    hex_bob = to_hex(keys_bob)

    # === AES-128 KEY MATCHING ===
    # Memilih Universal Hash yang nilai SHA-128 nya cocok antara Alice dan Bob
    aes_keys = []
    for idx in range(min(len(sha_keys_alice), len(sha_keys_bob))):
        if sha_keys_alice[idx] == sha_keys_bob[idx]:
            aes_keys.append(hex_alice[idx])

    end = time.time()
    time_hash = end - start
    
    return hex_alice, hex_bob, aes_keys, time_hash
