import cv2
from simple_facerec import SimpleFacerec
import pyttsx3
import re
import pandas as pd
import time

# Initialize the text-to-speech engine
engine = pyttsx3.init()

# Initialize the face recognizer
sfr = SimpleFacerec()
sfr.load_encoding_images("images/")

# Try different camera index if 2 doesn't work
cap = cv2.VideoCapture(0)  # Use 0 or 1 if 2 doesn't work

if not cap.isOpened():
    print("Error: Could not open camera.")
    exit()

# Initialize the attendance log (Excel sheet)
attendance_data = []
attendance_df = pd.DataFrame(columns=["Name", "In Time", "Out Time", "Status"])

# To store the last detection time for each person
last_detection_time = {}

# Minimum time (in seconds) for marking attendance as "Present"
min_time_for_present = 650

while True:
    ret, frame = cap.read()

    if not ret or frame is None:

        print("Failed to grab frame.")
        break

    # Detect Faces
    face_locations, face_names = sfr.detect_known_faces(frame)
    cleaned_names = []

    for name in face_names:
        # Clean the name using regex to remove anything after and including '('
        name_cleaned = re.sub(r"\(.*", "", name).strip()

        # Add cleaned name to the list for further checking
        cleaned_names.append(name_cleaned)

    # Remove duplicates and check if the names are already spoken
    cleaned_names = list(set(cleaned_names))

    # Get the current time
    current_time = time.time()

    # Check for each cleaned name
    for name_cleaned in cleaned_names:
        if name_cleaned not in last_detection_time:
            # Log the in-time when first detected
            last_detection_time[name_cleaned] = current_time
            # Play the greeting message
            engine.say(f"Hello {name_cleaned}, have a nice day.")
            engine.runAndWait()
            print(f"Detected: {name_cleaned}, In-Time: {current_time}")

        else:
            # Ensure we only calculate time_diff if last_detection_time is not None
            if last_detection_time[name_cleaned] is not None:
                # Calculate the time difference
                time_diff = current_time - last_detection_time[name_cleaned]

                if time_diff >= min_time_for_present:
                    # Mark "Present" if detected for more than 5 seconds
                    status = "Present"
                    # Log the out-time
                    out_time = current_time
                    print(f"Detected again: {name_cleaned}, Out-Time: {out_time}")
                    attendance_data.append([name_cleaned, time.strftime("%H:%M:%S", time.localtime(last_detection_time[name_cleaned])),
                                            time.strftime("%H:%M:%S", time.localtime(out_time)), status])
                    last_detection_time[name_cleaned] = None  # Reset last detection time for this person
                else:
                    # Mark "Absent" if detected for less than 5 seconds
                    status = "Absent"
                    out_time = current_time
                    print(f"Detected: {name_cleaned}, Out-Time: {out_time}, Status: {status}")
                    attendance_data.append([name_cleaned, time.strftime("%H:%M:%S", time.localtime(last_detection_time[name_cleaned])),
                                            time.strftime("%H:%M:%S", time.localtime(out_time)), status])

                    last_detection_time[name_cleaned] = None  # Reset last detection time for this person

    # If the attendance data has been updated, save it to Excel
    if attendance_data:
        attendance_df = pd.DataFrame(attendance_data, columns=["Name", "In Time", "Out Time", "Status"])
        attendance_df.to_excel("attendance_log.xlsx", index=False)

    # Add any additional processing logic here, if necessary

cap.release()
cv2.destroyAllWindows()
