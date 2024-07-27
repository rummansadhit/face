import cv2
import time
import logging
import subprocess

def setup_logging():
    logger = logging.getLogger('FaceDetectionBackgroundScript')
    handler = logging.FileHandler('C:\\FaceDetectionBackgroundScript.log')
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)
    return logger

def lock_all_sessions(logger):
    logger.info('Locking all sessions.')
    try:
        script_path = "C:\\temp\\face\\lock_screen.ps1"
        subprocess.Popen(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", script_path])
        logger.info('Lock all sessions script executed successfully.')
    except Exception as e:
        logger.error('Failed to execute lock all sessions script: %s', str(e))

def main():
    logger = setup_logging()
    logger.info('Starting main loop.')

    face_cascade = cv2.CascadeClassifier(cv2.data.haarcascades + 'haarcascade_frontalface_default.xml')
    cap = cv2.VideoCapture(0)

    if not cap.isOpened():
        logger.error('Failed to open the camera.')
        return

    no_face_count = 0
    try:
        while True:
            ret, frame = cap.read()
            if not ret:
                logger.warning('Failed to capture frame from camera.')
                continue

            gray = cv2.cvtColor(frame, cv2.COLOR_BGR2GRAY)
            faces = face_cascade.detectMultiScale(gray, scaleFactor=1.1, minNeighbors=5, minSize=(30, 30))

            if len(faces) == 0:
                no_face_count += 1
                logger.info('No face detected. Count: %d', no_face_count)
            else:
                no_face_count = 0
                logger.info('Face detected.')

            if no_face_count >= 50:  # No face detected for a certain number of frames
                logger.info('No face detected for sufficient frames. Locking all sessions.')
                lock_all_sessions(logger)
                no_face_count = 0

            time.sleep(0.1)  # Small delay to reduce CPU usage
    except KeyboardInterrupt:
        logger.info('Script terminated by user.')

    cap.release()
    cv2.destroyAllWindows()
    logger.info('Main loop exited.')

if __name__ == "__main__":
    main()