import numpy as np
import time
import csv
import math
import hashlib
import binascii
import socket
import pyaes
from tempfile import TemporaryFile
from random import seed, randint
from math import floor, pow, log
from PIL import Image
import subprocess
import sys
import matplotlib.pyplot as plt
import pandas as pd

# =========================================================================
# ============================== UNIVERSAL HASH ===========================
print("\n================================= UNIVERSAL HASH ===============================\n")

start5 = time.time()

# === Ubah dari XLS ke CSV ===
# Membaca file hasil BCH dari CSV, bukan Excel
with open(r"C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilBCH_eve/Bob_BCH_after_correction.csv", newline='') as f:
    reader = csv.reader(f)
    next(reader)  # lewati header jika ada
    aliice = [float(row[0]) for row in reader if row]

print("Panjang Input UnivHASH ALICE %d" % len(aliice))

key = []
jumlahkeya = math.floor(len(aliice) / 128)
jumlahkey = jumlahkeya

ukuranhash = 128
aaaa = len(aliice) % 128
lenalice = len(aliice) - aaaa

for i in range(0, aaaa):
    del aliice[lenalice]

Hashtab = []
with open('Hashtable128.csv', newline='') as f:
    reader = csv.reader(f)
    for row in reader:
        Hashtab.append(row)
Hashtable = np.array(Hashtab, dtype=int)

key1 = []
for i in range(jumlahkey):
    aaa = aliice[(i * ukuranhash): (ukuranhash * (i + 1))]
    print(
        f"Key-{i + 1}: Panjang data adalah {len(aaa)} dan ukuran HashTable yaitu {ukuranhash} x {len(aaa)}"
    )

    mat1 = []
    for x in range(len(Hashtable)):
        total = 0
        for y in range(len(aaa)):
            total += Hashtable[x][y] * aaa[y]
        mat1.append(int(total % 2))
    key1.append(mat1)
    print("Jumlah KEY ALICE Sekarang : ", len(key1))

v = 0
ax = [0] * 128 * jumlahkey
for i in range(jumlahkey):
    for j in range(128):
        ax[v] = key1[i][j]
        v += 1

univ = [ax[i] for i in range(len(ax))]

univ1 = np.array(univ).reshape(len(ax), 1)

# Simpan hasil Universal Hash ke CSV
with open(r"C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilHASH_eve/BOB_HASH.csv", "w", newline="") as fp:
    writer = csv.writer(fp)
    writer.writerow(["Alice"])
    for val in univ1:
        writer.writerow(val)

print("Universal Hash Alice disimpan ke files/BOB_HASH.csv")

end5 = time.time()
print('UNIVHASH Panjang bit hasil Universal Hash alice = %d, bob= %d' % (len(ax), len(ax)))
print('UNIVHASH Jumlah hasil key yang dibangkitkan = %d' % jumlahkey)
print('Waktu Proses HASHING : ', end5 - start5)

# =========================================================================
# ============================== CEK NIST =================================
print('\n\n====================== CEK NIST ========================')
print('============================================================\n')

startnist = time.time()
# command = "./NIST-TestALICE128"
# subprocess.Popen(command, shell=True)

indeks = []
indek = []

time.sleep(1)

with open(r"C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilNIST_eve/BOB_NIST.csv", newline='') as f:
    reader = csv.reader(f)
    for row in reader:
        indeks.append(row)

prindek = []
for i in range(len(indeks)):
    indek.append(int(indeks[i][0]))
    prindek.append(int(indeks[i][0]) + 1)

endnist = time.time()
print('===========~~~~~~~~~~~=========~~~~~~~~~~~~~~==============\n')
print('NIST Hasil prioritas index', prindek)
print('Waktu Proses Uji NIST : ', endnist - startnist)

# =========================================================================
# ================================= SHA-128 ===============================
print('\n\n====================== SHA-128 ========================')
print('============================================================\n')

start6 = time.time()
dataalice = []
hex1 = []
abc1 = []

# Check the length of 'key1[0]' to ensure valid indexing
key1_length = len(key1[0])
print(f"Length of key1[0]: {key1_length}")

# Ensure 'indek' has valid values and does not exceed the range of 'key1[0]'
valid_indeks = [idx for idx in indek if idx < key1_length]

# Initialize hex1 list by iterating over valid indices
for j in range(len(valid_indeks)):
    hash1 = [key1[0][valid_indeks[j]]] * 128
    data1 = "".join(str(e) for e in hash1)
    someText1 = data1.encode("ascii")
    b1 = hashlib.sha1(someText1).hexdigest()
    hex1.append(b1)
    print(f'Hasil hash Alice Kunci-{valid_indeks[j] + 1} = {hex1[j]}')

# === Simpan hasil hash ke CSV ===
csv_filename = 'C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilSHA_eve/BOB_SHA128.csv'
with open(csv_filename, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(['Alice'])
    for h in hex1:
        writer.writerow([h])
print(f"Data hash tersimpan di {csv_filename}")

# === Baca kembali CSV untuk verifikasi ===
hex2 = []
with open(csv_filename, mode='r') as file:
    reader = csv.reader(file)
    next(reader)
    for row in reader:
        if row:
            hex2.append(row[0])

# Verifikasi hasil hash
for j in range(len(valid_indeks)):
    if hex1[j] == hex2[j]:
        print(f'Hash Value-{valid_indeks[j] + 1} valid, proses enkripsi dapat dilakukan')
    else:
        print(f'Hash Value-{valid_indeks[j] + 1} not valid')

end6 = time.time()
print('Waktu Proses HASHING : ', end6 - start6)

import os
import binascii
import time

# =========================================================================
# ================================= AES-128 ===============================
print('\n\n===========~~~~~~~~~~~== AES ~~~~~~~~~~~~~~==============')
print('===========~~~~~~~~~~~=========~~~~~~~~~~~~~~==============\n')
start7 = time.time()

key_found = False
for kuncinya in range(min(len(indek), len(hex1))):
    if hex1[kuncinya] == hex2[kuncinya]:
        idx = indek[kuncinya]
        if idx >= len(key1):
            continue  # skip indeks yang tidak ada di key1
        keybit = ''.join(str(e) for e in key1[idx])
        keyint = int(keybit, 2)
        hex_str = '%x' % keyint
        if len(hex_str) % 2 != 0:
            hex_str = '0' + hex_str
        keybyte = binascii.unhexlify(hex_str)
        print(f'Key Bob 1 (16 bytes) = {keybyte}')
        key_found = True
        break

if not key_found:
    print("⚠️  Tidak ada pasangan hash yang cocok antara hex1 dan hex2!")
else:
    key_alice_1 = keybyte
    hex_key_alice_1 = key_alice_1.hex()
    print("Hexadecimal Key Alice:", hex_key_alice_1)

    output_dir = r"C:/Users/NADA VIDYAN SASMITA/OneDrive/Documents/semester 7/TA/skg/Output/P2P/hasilAES_evenew"
    os.makedirs(output_dir, exist_ok=True)
    output_file = os.path.join(output_dir, "BOB_AES.txt")

    with open(output_file, 'w', encoding='utf-8') as txt_file:
        txt_file.write(hex_key_alice_1)
    print(f"✅ Hex key saved successfully to {output_file}")

end7 = time.time()
print(f'Waktu proses AES-128: {end7 - start7:.4f} detik')

