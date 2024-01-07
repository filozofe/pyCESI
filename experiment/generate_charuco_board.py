#from Generating ArUco markers/opencv_generate_aruco.py
#useage: python opencv_generate_aruco.py --id 24 --type DICT_5X5_100 --output tags/DICT_5X5_100_id24.png
#older version of opencv-contrib-python
#pip install opencv-contrib-python==4.6.0.66


# import the necessary packages
import numpy as np
import argparse
import cv2
import sys

import cv2
from cv2 import aruco

aruco_dict = aruco.Dictionary_get(aruco.DICT_5X5_100)
aruco_dict.bytesList=aruco_dict.bytesList[30:,:,:]
board = aruco.CharucoBoard_create(7, 5, 1, 0.5, aruco_dict)

imboard = board.draw((2000, 2000))
cv2.imwrite("chessboard1.png", imboard)