Face Recognition Attendance System
This project is a Face Recognition-based Attendance System built with OpenCV and OpenPyXL.
It uses your computer's webcam to detect faces, record attendance in an Excel file, and store captured face images.

📌 Features
Detect faces in real-time using OpenCV's Haar Cascade Classifier.

Add new faces and assign names dynamically.

Mark attendance with the exact date and time.

Store attendance records in an attendance.xlsx file.

Save captured face images in a face_images folder.

Keep track of the total number of present individuals for each day.

📂 Project Structure
bash
Copy
Edit
.
├── shu.py               # Main Python script
├── attendance.xlsx      # Generated attendance record file
├── face_images/         # Saved face images
🛠️ Requirements
Python 3.x

OpenCV

OpenPyXL

Install dependencies:

bash
Copy
Edit
pip install opencv-python openpyxl
▶️ Usage
Run the script:

bash
Copy
Edit
python shu.py
Controls:

Press a to add a new face. Enter the name when prompted.

Press q to quit the application.

Attendance:

The system creates a sheet for the current date in attendance.xlsx.

Attendance is automatically marked when a known face is detected.

📊 Output Example
Excel Record Example:

Name	Time
John Doe	10:15:30
Jane Smith	10:17:45
Total Present: 2	

Console Output Example:

rust
Copy
Edit
Attendance marked for John Doe at 10:15:30
Attendance marked for Jane Smith at 10:17:45
⚠️ Notes
Ensure good lighting for accurate face detection.

All captured images are saved in face_images/ for future reference.

If attendance.xlsx doesn't exist, it will be created automatically.
