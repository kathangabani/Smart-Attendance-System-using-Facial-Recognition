
# Smart Attendance System Using Facial Recognition

## Description
This project implements a Smart Attendance System using facial recognition technology. It leverages a webcam to capture student faces and registers their attendance in an Excel sheet. The system uses OpenCV and machine learning techniques to recognize faces and match them against stored images for accurate attendance marking.

## Setup & Installation

### Prerequisites
- Python 3.x
- OpenCV
- dlib
- pandas
- xlwt (for writing Excel files)

### Install Dependencies
First, clone the repository:

```bash
git clone https://github.com/kathangabani/Smart-Attendance-System-using-Facial-Recognition.git
cd webcam_face_recognition-master
```

Next, install the required dependencies using pip:

```bash
pip install opencv-python pandas xlwt dlib
```

### Run the System
1. **Face Dataset Creation**
   - The system needs a set of known faces for recognition. To add a new face, run the `recognition.py` script, which captures the face image from your webcam.
   - The images will be stored in the `faces` folder.

2. **Train the Model**
   - Once the faces are captured, the system uses them to train the facial recognition model.

3. **Start the Attendance System**
   - Run the `main.py` script to start the webcam and begin recognizing faces in real time.

```bash
python main.py
```

The system will detect faces from the webcam and check against the stored faces, marking attendance automatically. The attendance records will be saved in an Excel file named `Attendance.xlsx`.

## Code Overview

### Files
- **recognition.py**: This script captures the face from the webcam and stores it in the `faces` directory.
- **main.py**: The main script that runs the attendance system and logs the attendance in real-time.
- **excel.py**: Manages the creation and writing to the Excel sheet.
- **Attendance.xlsx**: The Excel file where the attendance is logged.
- **faces/**: Directory containing the images of recognized faces for identification.

### Core Functionality
1. **Face Detection**: The system uses OpenCV's Haar Cascades for face detection.
2. **Face Recognition**: dlib’s facial recognition model is used to match detected faces with stored faces.
3. **Attendance Logging**: Each recognized face is logged into an Excel sheet using pandas.

## Running the System

To start the system, run the following command in your terminal:

```bash
python main.py
```

Once the system is running, the webcam will capture faces, and it will compare them with the known faces. If a match is found, it will mark the student as present in the `Attendance.xlsx` file.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
