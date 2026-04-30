import json
import math
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from Crypto.Cipher import AES
from Crypto.Util.Padding import unpad

ENC_FILE = "encrypted/gambar.enc"
META_FILE = "encrypted/gambar.meta.json"
OUTPUT_FILE = "gambar_decrypted.jpg"

with open(META_FILE, "r") as f:
    metadata = json.load(f)

key_id = metadata["key_id"]
iv = bytes.fromhex(metadata["iv"])

with open(key_id, "r", encoding="utf-8") as f:
    key_hex = f.read().strip()

try:
    key = bytes.fromhex(key_hex)
except ValueError:
    raise ValueError("File kunci tidak berisi format hexadecimal yang valid.")

if len(key) != 16:
    raise ValueError(f"Key AES-128 harus tepat 16 byte, tetapi didapatkan {len(key)} byte.")

with open(ENC_FILE, "rb") as f:
    ciphertext = f.read()

cipher = AES.new(key, AES.MODE_CBC, iv)
plaintext = unpad(cipher.decrypt(ciphertext), AES.block_size)

with open(OUTPUT_FILE, "wb") as f:
    f.write(plaintext)

print(f"Kunci AES-128 yang digunakan: {key_hex}")
print("Dekripsi berhasil:", OUTPUT_FILE)

# ==========================================
# Visualisasi Analogi Proses Dekripsi
# ==========================================
try:
    fig, axes = plt.subplots(1, 3, figsize=(15, 5))
    fig.suptitle('Proses Dekripsi Gambar', fontsize=16)

    # 1. Gambar Terenkripsi (Visualisasi byte acak)
    byte_array = np.frombuffer(ciphertext, dtype=np.uint8)
    side = int(math.ceil(math.sqrt(len(byte_array))))
    padded_array = np.pad(byte_array, (0, side*side - len(byte_array)), mode='constant')
    img_encrypted = padded_array.reshape((side, side))
    axes[0].imshow(img_encrypted, cmap='gray')
    axes[0].set_title('Gambar Terenkripsi (Ciphertext)')
    axes[0].axis('off')

    # 2. Kunci (Visualisasi Teks)
    axes[1].text(0.5, 0.5, f"Kunci (AES-128):\n\n{key_hex}", 
                 fontsize=12, ha='center', va='center', wrap=True, family='monospace')
    axes[1].set_title('Kunci Dekripsi')
    axes[1].axis('off')

    # 3. Gambar Terdekripsi
    img_dekrip = mpimg.imread(OUTPUT_FILE)
    axes[2].imshow(img_dekrip)
    axes[2].set_title('Gambar Terdekripsi (Asli)')
    axes[2].axis('off')

    plt.tight_layout()
    plt.savefig('analogi_dekripsi.png')
    plt.show()
except Exception as e:
    print(f"Tidak dapat menampilkan plot visualisasi: {e}")