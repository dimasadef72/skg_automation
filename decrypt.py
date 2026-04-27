import json
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