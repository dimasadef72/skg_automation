import cv2
import numpy as np
from Crypto.Cipher import AES
from Crypto.Util.Padding import pad, unpad
import os
import binascii
import matplotlib.pyplot as plt

def read_key_from_file(filepath):
    try:
        with open(filepath, 'r') as f:
            key_hex = f.read().strip()
        # Jika file berisi "Tidak ada" atau kosong
        if key_hex.lower() == "tidak ada" or not key_hex:
            return None, key_hex
        return binascii.unhexlify(key_hex), key_hex
    except Exception as e:
        print(f"Error membaca kunci dari {filepath}: {e}")
        return None, ""

def encrypt_image(image, key):
    img_bytes = image.tobytes()
    padded_bytes = pad(img_bytes, AES.block_size)
    iv = os.urandom(16) # Inisialisasi Vector acak untuk CBC
    cipher = AES.new(key, AES.MODE_CBC, iv)
    ciphertext = cipher.encrypt(padded_bytes)
    return ciphertext, iv

def decrypt_image_to_bytes(ciphertext, key, iv):
    try:
        cipher = AES.new(key, AES.MODE_CBC, iv)
        padded_bytes = cipher.decrypt(ciphertext)
        img_bytes = unpad(padded_bytes, AES.block_size)
        return img_bytes
    except ValueError as e:
        # Menangkap error padding yang biasanya terjadi jika kunci salah
        return None
    except Exception as e:
        print(f"Terjadi kesalahan saat dekripsi: {e}")
        return None

def main():
    img_path = r"D:\skg_automation\hewan_hama\tikus.jpg"
    alice_bob_key_path = r"D:\skg_automation\Output_aul\skenario_2\kunci_aes128\BB1_kunci_alice_bob.txt"
    eve_key_path = r"D:\skg_automation\Output_aul\skenario_2\kunci_aes128\BB1_kunci_eve.txt"

    print("="*50)
    print("SIMULASI ENKRIPSI & DEKRIPSI GAMBAR DENGAN HASIL SKG")
    print("="*50)

    # 1. Membaca Gambar
    if not os.path.exists(img_path):
        print(f"ERROR: Gambar tidak ditemukan di {img_path}")
        return
    img = cv2.imread(img_path)
    if img is None:
        print(f"ERROR: Gagal memuat gambar. Pastikan format gambar valid.")
        return
        
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB) # Konversi BGR (OpenCV) ke RGB (Matplotlib)
    shape = img_rgb.shape
    dtype = img_rgb.dtype

    # 2. Membaca Kunci
    alice_key, alice_key_hex = read_key_from_file(alice_bob_key_path)
    eve_key, eve_key_hex = read_key_from_file(eve_key_path)

    if alice_key is None or len(alice_key) != 16:
        print("ERROR: Kunci Alice/Bob tidak valid. Harus 16 byte (128 bit) dari nilai Hex.")
        return

    print(f"[+] Kunci Alice & Bob (Hex) : {alice_key_hex}")
    
    if eve_key is None:
        print(f"[+] Kunci Eve               : {eve_key_hex} (Eve tidak memiliki kunci dari SKG)")
        print("    -> Menggunakan kunci acak (16 byte) sebagai simulasi upaya Eve untuk menebak.")
        eve_key = os.urandom(16) 
    else:
        print(f"[+] Kunci Eve (Hex)         : {eve_key_hex}")

    # 3. Proses Enkripsi oleh Alice
    print("\n[!] Alice sedang mengenkripsi gambar...")
    ciphertext, iv = encrypt_image(img_rgb, alice_key)
    
    # Visualisasi Ciphertext (memotong ciphertext sesuai ukuran gambar asli untuk ditampilkan)
    encrypted_bytes = bytearray(ciphertext)[:np.prod(shape)]
    encrypted_img = np.frombuffer(encrypted_bytes, dtype=dtype).reshape(shape)

    # 4. Proses Dekripsi oleh Bob (Penerima Sah)
    print("[!] Bob sedang mendekripsi gambar dengan kunci yang sah...")
    decrypted_bytes = decrypt_image_to_bytes(ciphertext, alice_key, iv)
    if decrypted_bytes:
        decrypted_img = np.frombuffer(decrypted_bytes, dtype=dtype).reshape(shape)
        bob_success = True
        print("    -> Bob BERHASIL mendekripsi gambar!")
    else:
        decrypted_img = np.zeros(shape, dtype=dtype)
        bob_success = False
        print("    -> Bob GAGAL mendekripsi gambar.")

    # 5. Proses Dekripsi oleh Eve (Penyadap)
    print("[!] Eve mencoba menyadap dan mendekripsi gambar dengan kuncinya...")
    eve_decrypted_bytes = decrypt_image_to_bytes(ciphertext, eve_key, iv)
    
    if eve_decrypted_bytes and len(eve_decrypted_bytes) == np.prod(shape):
        eve_img = np.frombuffer(eve_decrypted_bytes, dtype=dtype).reshape(shape)
        print("    -> PERINGATAN: Eve BERHASIL mendekripsi gambar!")
    else:
        print("    -> Eve GAGAL menyadap! (Error padding / Kunci tidak cocok).")
        # Visualisasi kegagalan Eve (Eve hanya melihat noise/sampah jika mencoba mendekripsi)
        try:
            cipher_eve = AES.new(eve_key, AES.MODE_CBC, iv)
            eve_raw = cipher_eve.decrypt(ciphertext)
            eve_img = np.frombuffer(eve_raw[:np.prod(shape)], dtype=dtype).reshape(shape)
        except:
            eve_img = np.random.randint(0, 256, shape, dtype=np.uint8)

    # 6. Menampilkan Hasil Visual dengan Matplotlib
    print("\n[+] Membuka jendela tampilan hasil...")
    plt.figure(figsize=(16, 5))
    plt.suptitle("Hasil Enkripsi dan Dekripsi Gambar (AES-128)", fontsize=16)
    
    plt.subplot(1, 4, 1)
    plt.title("1. Gambar Asli (Alice)")
    plt.imshow(img_rgb)
    plt.axis('off')

    plt.subplot(1, 4, 2)
    plt.title("2. Gambar Terenkripsi")
    plt.imshow(encrypted_img)
    plt.axis('off')

    plt.subplot(1, 4, 3)
    if bob_success:
        plt.title("3. Dekripsi Bob (Berhasil)")
        plt.imshow(decrypted_img)
    else:
        plt.title("3. Dekripsi Bob (Gagal)")
        plt.imshow(np.zeros_like(img_rgb))
    plt.axis('off')

    plt.subplot(1, 4, 4)
    plt.title("4. Dekripsi Eve (Gagal)")
    plt.imshow(eve_img)
    plt.axis('off')

    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    main()
