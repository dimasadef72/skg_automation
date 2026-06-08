import cv2
import numpy as np
from ultralytics import YOLO
import torch
import time
import os
from picamera2 import Picamera2

# =========================
# DEVICE (CPU/GPU otomatis)
# =========================
device = "cuda" if torch.cuda.is_available() else "cpu"
print(f"Using device: {device}")

# =========================
# LOAD MODEL
# =========================
script_dir = os.path.dirname(os.path.abspath(__file__))
model_path = os.path.join(script_dir, "best.pt")
model = YOLO(model_path).to(device)

try:
    print("Model class names:", model.names)
except Exception:
    pass


# =========================
# IMAGE MODE CHECK
# =========================
INPUT_SOURCE = 0
use_image = isinstance(INPUT_SOURCE, str)

if use_image:
    single_image = cv2.imread(INPUT_SOURCE)
    if single_image is None:
        raise FileNotFoundError(f"Gambar tidak ditemukan: {INPUT_SOURCE}")
    print(f"Using image input: {INPUT_SOURCE}")
    picam2 = None
else:
    print("Using Picamera2 camera (Raspberry Pi)")
    picam2 = Picamera2()

    config = picam2.create_preview_configuration(
        main={"size": (640, 480)}
    )

    picam2.configure(config)
    picam2.start()


# =========================
# CLASS FILTER (TETAP PUNYAMU)
# =========================
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

class_conf_override = {
    "Keong": 0.15,
    "Wereng": 0.15,
    "Wereng putih": 0.15,
    "Kresek": 0.15,
    "Ulat": 0.15,
}


# =========================
# LOOP INFERENCE
# =========================
while True:
    start_time = time.time()

    # =====================
    # AMBIL FRAME (FIX DI SINI)
    # =====================
    if use_image:
        frame = single_image.copy()
    else:
        frame = picam2.capture_array()

    # =====================
    # PREPROCESS
    # =====================
    img = cv2.resize(frame, (640, 480), interpolation=cv2.INTER_LINEAR)

    conf_thresh = 0.30
    iou_thresh = 0.45

    results = model.predict(
        source=img,
        device=device,
        conf=conf_thresh,
        iou=iou_thresh,
        imgsz=960,
        task="segment"
    )

    end_time = time.time()
    fps = 1 / (end_time - start_time)

    # =====================
    # DRAW RESULT (TETAP)
    # =====================
    for r in results:
        display = img.copy()
        min_area = 300

        try:
            boxes = r.boxes
            if boxes is not None and len(boxes) > 0:
                xyxy = boxes.xyxy.cpu().numpy()
                confs = boxes.conf.cpu().numpy()
                clss = boxes.cls.cpu().numpy().astype(int)

                raw_summary = []

                for box, conf, cls_id in zip(xyxy, confs, clss):
                    x1, y1, x2, y2 = box.astype(int)
                    area = max(0, (x2 - x1)) * max(0, (y2 - y1))

                    name = model.names.get(cls_id, str(cls_id))

                    raw_summary.append(f"{name}:{conf:.2f}")

                    per_class_conf = class_conf_override.get(name, conf_thresh)

                    if conf < per_class_conf:
                        continue
                    if area < min_area:
                        continue
                    if name not in allowed_classes:
                        continue

                    cv2.rectangle(display, (x1, y1), (x2, y2), (0, 255, 0), 2)
                    cv2.putText(
                        display,
                        f"{name} {conf:.2f}",
                        (x1, max(15, y1 - 5)),
                        cv2.FONT_HERSHEY_SIMPLEX,
                        0.6,
                        (0, 255, 0),
                        2
                    )

                if raw_summary:
                    print("Raw detections:", ", ".join(raw_summary))

        except Exception:
            display = r.plot()

        cv2.putText(
            display,
            f"FPS: {fps:.2f}",
            (10, 30),
            cv2.FONT_HERSHEY_SIMPLEX,
            1,
            (0, 0, 255),
            2
        )

        cv2.imshow("YOLOv11 - Segmentation", display)

    # =====================
    # EXIT
    # =====================
    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

    if use_image:
        break


# =========================
# CLEANUP
# =========================
if picam2 is not None:
    picam2.stop()

cv2.destroyAllWindows()