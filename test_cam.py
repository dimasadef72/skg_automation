import cv2
cap = cv2.VideoCapture(0, cv2.CAP_V4L2)

cap.set(cv2.CAP_PROP_FRAME_WIDTH, 640)
cap.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)

print("opened =", cap.isOpened())

ret, frame = cap.read()
print("ret =", ret)

if ret:
    print(frame.shape)
    cv2.imwrite("test.jpg", frame)
    print("saved test.jpg")

cap.release()