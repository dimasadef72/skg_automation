import cv2
import numpy as np
from ultralytics import YOLO
import torch
import time
import os

# Pastikan CUDA tersedia
device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Using device: {device}")

# Load model YOLOv11 ke GPU jika tersedia
script_dir = os.path.dirname(os.path.abspath(__file__))
model_path = os.path.join(script_dir, "best.pt")
model = YOLO(model_path).to(device)
# Tampilkan daftar nama kelas agar bisa disesuaikan dengan whitelist
try:
    print("Model class names:", model.names)
except Exception:
    pass

# Ubah ke path file gambar kalau ingin inference dari gambar lokal.
# Contoh: INPUT_SOURCE = r"C:\Users\Nama\Pictures\keong.jpg"
INPUT_SOURCE = 0


def open_camera_source(source_index):
    # Windows sering lebih stabil dengan DirectShow daripada MSMF.
    for backend in (cv2.CAP_DSHOW, cv2.CAP_MSMF):
        cap_obj = cv2.VideoCapture(source_index, backend)
        if cap_obj.isOpened():
            print(f"Camera opened with backend: {backend}")
            return cap_obj
        cap_obj.release()
    return cv2.VideoCapture(source_index)

# Default whitelist sesuai kelas yang memang ada di model ini
allowed_classes = [
    "Belalang",
    "Bercak putih",
    "Blast",
    "Capung",
    "Keong",
    "Kepik",
    "Kresek",
    "Ulat",
    "Wereng",
    "Wereng putih",
]

# Threshold yang bisa dibuat lebih longgar untuk kelas yang sering miss
class_conf_override = {
    "Keong": 0.15,
    "Wereng": 0.15,
    "Wereng putih": 0.15,
    "Kresek": 0.15,
    "Ulat": 0.15,
}

use_image = isinstance(INPUT_SOURCE, str)
if use_image:
    single_image = cv2.imread(INPUT_SOURCE)
    if single_image is None:
        raise FileNotFoundError(f"Gambar tidak ditemukan atau gagal dibuka: {INPUT_SOURCE}")
    print(f"Using image input: {INPUT_SOURCE}")
    cap = None
else:
    print(f"Using camera input: {INPUT_SOURCE}")
    cap = open_camera_source(INPUT_SOURCE)
    if not cap.isOpened():
        raise RuntimeError(
            f"Kamera tidak bisa dibuka. Coba ganti INPUT_SOURCE ke 1/2, atau pakai file gambar."
        )

while True:
    start_time = time.time()
    if use_image:
        frame = single_image.copy()
        ret = True
    else:
        ret, frame = cap.read()
        if not ret:
            break

    # Resize gambar agar sesuai dengan input model (harus kelipatan 32)
    img = cv2.resize(frame, (640, 480), interpolation=cv2.INTER_LINEAR)

    # Gunakan array langsung dan lakukan inference dengan ambang kepercayaan dan IOU
    # Turunkan threshold untuk membantu objek kecil seperti keong tetap terdeteksi
    conf_thresh = 0.30
    iou_thresh = 0.45
    results = model.predict(source=img, device=device, conf=conf_thresh, iou=iou_thresh, imgsz=960, task="segment")

    # Hitung FPS
    end_time = time.time()
    fps = 1 / (end_time - start_time)

    # Tampilkan hasil deteksi dengan filtering sederhana
    for r in results:
        display = img.copy()

        # Untuk objek kecil, jangan terlalu besar agar keong tetap lolos.
        min_area = 300  # minimal area (pixel) untuk mengabaikan deteksi noise kecil

        try:
            boxes = r.boxes
            if boxes is not None and len(boxes) > 0:
                xyxy = boxes.xyxy.cpu().numpy()
                confs = boxes.conf.cpu().numpy()
                clss = boxes.cls.cpu().numpy().astype(int)

                # Debug ringan: lihat deteksi mentah sebelum filter
                raw_summary = []

                for box, conf, cls_id in zip(xyxy, confs, clss):
                    x1, y1, x2, y2 = box.astype(int)
                    area = max(0, (x2 - x1)) * max(0, (y2 - y1))
                    name = model.names.get(int(cls_id), str(int(cls_id))) if hasattr(model, 'names') else str(int(cls_id))

                    raw_summary.append(f"{name}:{conf:.2f}")

                    per_class_conf = class_conf_override.get(name, conf_thresh)

                    if conf < per_class_conf:
                        continue
                    if area < min_area:
                        continue
                    if allowed_classes is not None and name not in allowed_classes:
                        continue

                    color = (0, 255, 0)
                    cv2.rectangle(display, (x1, y1), (x2, y2), color, 2)
                    cv2.putText(display, f'{name} {conf:.2f}', (x1, max(15, y1 - 5)),
                                cv2.FONT_HERSHEY_SIMPLEX, 0.6, color, 2)

                if raw_summary:
                    print("Raw detections:", ", ".join(raw_summary))
        except Exception:
            # Jika struktur hasil berbeda, fallback ke plot()
            try:
                display = r.plot()
            except Exception:
                display = img.copy()

        cv2.putText(display, f'FPS: {fps:.2f}', (10, 30),
                    cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
        
        cv2.imshow("YOLOv11 - Segmentation", display)

    # Tekan 'q' untuk keluar
    wait_key_delay = 0 if use_image else 1
    if cv2.waitKey(wait_key_delay) & 0xFF == ord('q'):
        break

    if use_image:
        break

if cap is not None:
    cap.release()
cv2.destroyAllWindows()