from sklearn.neighbors import KNeighborsClassifier
import cv2
import pickle
import numpy as np
import os
import csv
import time
from datetime import datetime
from win32com.client import Dispatch

def speak(str1):
    speaker = Dispatch("SAPI.SpVoice")
    speaker.Speak(str1)

# Open webcam
video = cv2.VideoCapture(0)

# Load Haar Cascade for face detection
facedetect = cv2.CascadeClassifier('data/haarcascade_frontalface_default.xml')

# Load training data
with open('data/names.pkl', 'rb') as w:
    LABELS = pickle.load(w)
with open('data/faces_data.pkl', 'rb') as f:
    FACES = pickle.load(f)

print('Shape of Faces matrix --> ', FACES.shape)  # Should be (100, 2175)

# Train KNN classifier
knn = KNeighborsClassifier(n_neighbors=5)
knn.fit(FACES, LABELS)

# Load background image
imgBackground = cv2.imread("background.png")

# CSV header
COL_NAMES = ['NAME', 'TIME']

while True:
    ret, frame = video.read()
    gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
    faces = facedetect.detectMultiScale(gray, 1.3, 5)

    for (x, y, w, h) in faces:
        crop_img = frame[y:y+h, x:x+w]

        # ✅ Convert cropped face to grayscale
        gray_face = cv2.cvtColor(crop_img, cv2.COLOR_BGR2GRAY)

        # ✅ Resize to 29x75 to match training input size (29*75 = 2175)
        resized_img = cv2.resize(gray_face, (29, 75))  # (width, height)
        resized_img = resized_img.flatten().reshape(1, -1)  # Shape (1, 2175)

        # Predict
        output = knn.predict(resized_img)

        # Time & Date
        ts = time.time()
        date = datetime.fromtimestamp(ts).strftime("%d-%m-%Y")
        timestamp = datetime.fromtimestamp(ts).strftime("%H:%M:%S")
        attendance_file = f"Attendance/Attendance_{date}.csv"
        file_exists = os.path.isfile(attendance_file)

        # Draw bounding box & name
        cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 2)
        cv2.rectangle(frame, (x, y - 40), (x + w, y), (50, 50, 255), -1)
        cv2.putText(frame, str(output[0]), (x, y - 10),
                    cv2.FONT_HERSHEY_COMPLEX, 1, (255, 255, 255), 1)

        attendance = [str(output[0]), str(timestamp)]

    # Show final frame inside background
    imgBackground[162:162 + 480, 55:55 + 640] = frame
    cv2.imshow("Frame", imgBackground)

    k = cv2.waitKey(1)
    if k == ord('o'):
        speak("Attendance Taken.")
        time.sleep(2)

        with open(attendance_file, 'a', newline='') as csvfile:
            writer = csv.writer(csvfile)
            if not file_exists:
                writer.writerow(COL_NAMES)
            writer.writerow(attendance)

    if k == ord('q'):
        break

video.release()
cv2.destroyAllWindows()
