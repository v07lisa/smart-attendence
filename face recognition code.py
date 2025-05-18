import face_recognition
import cv2
import numpy as np
import os
import xlwt
from xlwt import Workbook
from datetime import date
import xlrd
from xlutils.copy import copy as xl_copy

# Get current working directory
CurrentFolder = os.getcwd()
image1_path = os.path.join(CurrentFolder, 'rahul.png')
image2_path = os.path.join(CurrentFolder, 'sneha.png')

# Initialize webcam
video_capture = cv2.VideoCapture(0)

if not video_capture.isOpened():
    print("Error: Could not open webcam.")
    exit()

# Load images and encode faces
person1_name = "Rahul"
person1_image = face_recognition.load_image_file(image1_path)
person1_face_encoding = face_recognition.face_encodings(person1_image)[0]

person2_name = "Sneha"
person2_image = face_recognition.load_image_file(image2_path)
person2_face_encoding = face_recognition.face_encodings(person2_image)[0]

# Store known encodings and names
known_face_encodings = [person1_face_encoding, person2_face_encoding]
known_face_names = [person1_name, person2_name]

# Variables for face recognition
face_locations = []
face_encodings = []
face_names = []
process_this_frame = True
attendance_taken = set()  # Keep track of recorded students

# Excel file setup
excel_filename = 'attendence_excel.xls'
if os.path.exists(excel_filename):
    rb = xlrd.open_workbook(excel_filename, formatting_info=True)
    wb = xl_copy(rb)
else:
    wb = Workbook()

subject = input("Enter subject name: ")
sheet1 = wb.add_sheet(subject)
sheet1.write(0, 0, 'Name/Date')
sheet1.write(0, 1, str(date.today()))
row = 1

# Start video capture
while True:
    ret, frame = video_capture.read()
    if not ret:
        print("Error: Failed to capture frame")
        break

    # Resize frame for faster processing
    small_frame = cv2.resize(frame, (0, 0), fx=0.25, fy=0.25)
    rgb_small_frame = small_frame[:, :, ::-1]  # Convert BGR to RGB

    if process_this_frame:
        # Detect faces and get encodings
        face_locations = face_recognition.face_locations(rgb_small_frame)
        face_encodings = face_recognition.face_encodings(rgb_small_frame, face_locations)

        face_names = []
        for face_encoding in face_encodings:
            # Compare face with known faces
            matches = face_recognition.compare_faces(known_face_encodings, face_encoding)
            face_distances = face_recognition.face_distance(known_face_encodings, face_encoding)
            best_match_index = np.argmin(face_distances)

            if matches[best_match_index]:
                name = known_face_names[best_match_index]
            else:
                name = "Unknown"

            face_names.append(name)

            # Mark attendance if not already taken
            if name != "Unknown" and name not in attendance_taken:
                sheet1.write(row, 0, name)
                sheet1.write(row, 1, "Present")
                row += 1
                attendance_taken.add(name)
                wb.save(excel_filename)
                print(f"Attendance taken for {name}")

    process_this_frame = not process_this_frame

    # Display results
    for (top, right, bottom, left), name in zip(face_locations, face_names):
        top, right, bottom, left = top * 4, right * 4, bottom * 4, left * 4

        # Draw face box
        cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)

        # Label face
        cv2.rectangle(frame, (left, bottom - 35), (right, bottom), (0, 0, 255), cv2.FILLED)
        font = cv2.FONT_HERSHEY_DUPLEX
        cv2.putText(frame, name, (left + 6, bottom - 6), font, 1.0, (255, 255, 255), 1)

    # Show video
    cv2.imshow('Video', frame)

    # Quit on 'q' press
    if cv2.waitKey(1) & 0xFF == ord('q'):
        print("Saving data and closing...")
        break

# Cleanup
video_capture.release()
cv2.destroyAllWindows()
