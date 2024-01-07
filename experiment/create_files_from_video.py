# inspired from https://mecaruco2.readthedocs.io/en/latest/notebooks_rst/Aruco/sandbox/ludovic/aruco_calibration_rotation.html

import cv2
import os

def save_all_frames(video_path, dir_path, basename, ext='jpg'):
    cap = cv2.VideoCapture(video_path)

    if not cap.isOpened():
        return

    os.makedirs(dir_path, exist_ok=True)
    base_path = os.path.join(dir_path, basename)

    digit = len(str(int(cap.get(cv2.CAP_PROP_FRAME_COUNT))))

    n = 0
    i=0

    while True:
        ret, frame = cap.read()
        if ret:
            if n % 15 == 0:
                cv2.imwrite('{}_{}.{}'.format(base_path, str(i).zfill(4), ext), frame)
                i += 1
            n += 1
        else:
            return

save_all_frames('http://192.168.1.18:4747/mjpegfeed?640x480', './data/', 'sample_video_img', 'png')
