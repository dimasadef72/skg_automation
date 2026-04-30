import json
import os
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad
from Crypto.Random import get_random_bytes

KEY_ID = "Output_aul/skenario_1/kunci_aes128/BB100_kunci_alice_bob.txt"
INPUT_FILE = "tikus.jpg"
OUTPUT_FILE = "encrypted/gambar.enc"
META_FILE = "encrypted/gambar.meta.json"

os.makedirs("encrypted", exist_ok=True)

with open(KEY_ID, "r", encoding="utf-8") as f:
    key_hex = f.read().strip()

try:
    key = bytes.fromhex(key_hex)
except ValueError:
    raise ValueError("File kunci tidak berisi format hexadecimal yang valid.")

if len(key) != 16:
    raise ValueError(f"Key AES-128 harus tepat 16 byte, tetapi didapatkan {len(key)} byte.")

iv = get_random_bytes(16)
cipher = AES.new(key, AES.MODE_CBC, iv)

with open(INPUT_FILE, "rb") as f:
    plaintext = f.read()

ciphertext = cipher.encrypt(pad(plaintext, AES.block_size))

with open(OUTPUT_FILE, "wb") as f:
    f.write(ciphertext)

metadata = {
    "algorithm": "AES-128-CBC",
    "key_id": KEY_ID,
    "iv": iv.hex(),
    "original_name": INPUT_FILE
}

with open(META_FILE, "w") as f:
    json.dump(metadata, f, indent=4)

print(f"Kunci AES-128 yang digunakan: {key_hex}")
print("Enkripsi berhasil:", OUTPUT_FILE)

# ==========================================
# Visualisasi Analogi Proses Enkripsi
# ==========================================
try:
    fig, axes = plt.subplots(1, 3, figsize=(15, 5))
    fig.suptitle('Proses Enkripsi Gambar', fontsize=16)

    # 1. Gambar Asli
    img_asli = mpimg.imread(INPUT_FILE)
    axes[0].imshow(img_asli)
    axes[0].set_title('Gambar Asli (Tikus)')
    axes[0].axis('off')

    # 2. Kunci (Visualisasi Teks)
    axes[1].text(0.5, 0.5, f"Kunci (AES-128):\n\n{key_hex}", 
                 fontsize=12, ha='center', va='center', wrap=True, family='monospace')
    axes[1].set_title('Kunci Enkripsi')
    axes[1].axis('off')

    # 3. Gambar Terenkripsi (Visualisasi byte acak)
    byte_array = np.frombuffer(ciphertext, dtype=np.uint8)
    side = int(math.ceil(math.sqrt(len(byte_array))))
    padded_array = np.pad(byte_array, (0, side*side - len(byte_array)), mode='constant')
    img_encrypted = padded_array.reshape((side, side))
    axes[2].imshow(img_encrypted, cmap='gray')
    axes[2].set_title('Gambar Terenkripsi (Ciphertext)')
    axes[2].axis('off')

    plt.tight_layout()
    plt.savefig('analogi_enkripsi.png')
    plt.show()
except Exception as e:
    print(f"Tidak dapat menampilkan plot visualisasi: {e}")