import cv2
import os

haar_file = 'haarcascade_frontalface_default.xml'
datasets = 'datasets'

sub_data = str(input("Input Student Data: "))

MAX_IMAGES_COUNT = 120


def start_training(max_count):
    (width, height) = (130, 100)  # defining the size of images
    face_cascade = cv2.CascadeClassifier(haar_file)
    webcam = cv2.VideoCapture(0)
    count = 0
    while count <= max_count:
        (_, im) = webcam.read()
        gray = cv2.cvtColor(im, cv2.COLOR_BGR2GRAY)
        faces = face_cascade.detectMultiScale(gray, 1.3, 4)
        for (x, y, w, h) in faces:
            cv2.rectangle(im, (x, y), (x + w, y + h), (255, 0, 0), 2)
            cv2.putText(im, f"Wait - {max_count-count} Images Remaining", (x + 10, y - 10),
                        cv2.FONT_HERSHEY_PLAIN, 1, (0, 255, 0))
            face = gray[y:y + h, x:x + w]
            face_resize = cv2.resize(face, (width, height))
            cv2.imwrite('%s/%s.png' % (path, count), face_resize)
            count += 1

        cv2.imshow('OpenCV', im)
        key = cv2.waitKey(10)
        if key == 27:
            break


path = os.path.join(datasets, sub_data)
if not os.path.isdir(path):
    os.mkdir(path)

start_training(MAX_IMAGES_COUNT)
