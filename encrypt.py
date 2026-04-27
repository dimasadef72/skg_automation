import json
import os
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