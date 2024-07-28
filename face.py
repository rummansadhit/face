import cv2
import time
import logging
import subprocess
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from PIL import Image, ImageTk, ImageDraw, ImageFont
import numpy as np
import face_recognition
import win32com.client

class FaceDetectionApp:
    def __init__(self, root):
        self.root = root
        self.logger = self.setup_logging()
        self.cap = None
        self.running = False
        self.camera_indices = []
        self.camera_names = []

        self.setup_gui()

    def setup_logging(self):
        logger = logging.getLogger('FaceDetectionBackgroundScript')
        handler = logging.FileHandler('C:\\FaceDetectionBackgroundScript.log')
        formatter = logging.Formatter('%(asctime)s - %(name)s - %(levellevel)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.setLevel(logging.INFO)
        return logger

    def lock_all_sessions(self):
        self.logger.info('Locking all sessions.')
        try:
            script_path = "C:\\temp\\face\\lock_screen.ps1"
            subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path])
            self.logger.info('Lock all sessions script executed successfully.')
        except Exception as e:
            self.logger.error('Failed to execute lock all sessions script: %s', str(e))

    def list_cameras(self):
        # List cameras using OpenCV and match with WMI names
        self.camera_indices = []
        self.camera_names = []
        index = 0

        # Get camera names using WMI
        wmi = win32com.client.GetObject("winmgmts:")
        cameras = wmi.InstancesOf("Win32_PnPEntity")
        wmi_camera_names = []
        for camera in cameras:
            if camera.Caption and ("camera" in camera.Caption.lower() or "video" in camera.Caption.lower()):
                wmi_camera_names.append(camera.Caption)

        while True:
            cap = cv2.VideoCapture(index)
            if not cap.read()[0]:
                break
            self.camera_indices.append(index)
            # Add a default name if WMI name is not found
            camera_name = wmi_camera_names[index] if index < len(wmi_camera_names) else f"Camera {index}"
            self.camera_names.append(camera_name)
            cap.release()
            index += 1

    def start_face_detection(self, camera_index, lock_delay):
        self.cap = cv2.VideoCapture(camera_index)
        if not self.cap.isOpened():
            self.logger.error('Failed to open the camera.')
            messagebox.showerror("Error", "Failed to open the camera.")
            return

        self.no_face_count = 0
        self.lock_delay = lock_delay
        self.frames_to_lock = int(self.lock_delay * 10)  # converting seconds to frames (assuming 10 fps)
        self.running = True
        self.update_frame()

    def update_frame(self):
        if not self.running:
            return

        ret, frame = self.cap.read()
        if not ret:
            self.logger.warning('Failed to capture frame from camera.')
            self.root.after(10, self.update_frame)
            return

        # Use face_recognition to detect faces
        rgb_frame = frame[:, :, ::-1]  # Convert BGR to RGB
        face_locations = face_recognition.face_locations(rgb_frame)

        if len(face_locations) == 0:
            self.no_face_count += 1
            self.logger.info('No face detected. Count: %d', self.no_face_count)
        else:
            self.no_face_count = 0
            self.logger.info('Face detected.')
            # Draw "Face Detected" text on the frame
            frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            img = Image.fromarray(frame)
            draw = ImageDraw.Draw(img)
            font = ImageFont.truetype("arial.ttf", 32)  # Adjust the font size as needed
            draw.text((10, 10), "Face Detected", font=font, fill=(255, 0, 0))
            frame = np.array(img)  # Convert back to OpenCV format

        if self.no_face_count >= self.frames_to_lock:  # No face detected for a certain number of frames
            self.logger.info('No face detected for sufficient frames. Locking all sessions.')
            self.lock_all_sessions()
            self.no_face_count = 0

        # Convert the frame to ImageTk format
        frame = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        img = Image.fromarray(frame)
        imgtk = ImageTk.PhotoImage(image=img)

        # Update the image panel
        self.panel.imgtk = imgtk
        self.panel.config(image=imgtk)

        self.root.after(10, self.update_frame)

    def on_start_button_click(self):
        selected_index = self.camera_selection.curselection()
        if not selected_index:
            messagebox.showerror("Error", "Please select a camera.")
            return
        camera_index = self.camera_indices[selected_index[0]]
        lock_delay = int(self.lock_delay_combobox.get())
        self.start_face_detection(camera_index, lock_delay)

    def on_stop_button_click(self):
        self.running = False
        if self.cap:
            self.cap.release()
            self.cap = None
        self.panel.config(image='')

    def setup_gui(self):
        self.root.title("Face Detection Camera Selector")

        tk.Label(self.root, text="Select a Camera:").pack()

        self.camera_selection = tk.Listbox(self.root)
        self.list_cameras()
        for cam in self.camera_names:
            self.camera_selection.insert(tk.END, cam)
        self.camera_selection.pack()

        tk.Label(self.root, text="Select Lock Delay (seconds):").pack()

        self.lock_delay_combobox = ttk.Combobox(self.root, values=[str(i) for i in range(1, 11)])
        self.lock_delay_combobox.current(4)  # Default to 5 seconds
        self.lock_delay_combobox.pack()

        start_button = tk.Button(self.root, text="Start", command=self.on_start_button_click)
        start_button.pack()

        stop_button = tk.Button(self.root, text="Stop", command=self.on_stop_button_click)
        stop_button.pack()

        self.panel = tk.Label(self.root)
        self.panel.pack()

def main():
    root = tk.Tk()
    app = FaceDetectionApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
