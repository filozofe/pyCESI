# This is a sample Python script.

# Press Maj+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.



# ouvrir espcam avec OpenCV
""" URL = "insert here the Esp-cam URL, eg. http://192.168.1.1"
AWB = True
cap = cv2.VideoCapture(URL + ":81/stream") """


import numpy as np
import cv2 as cv
import datetime


cap = cv.VideoCapture(0)
fps = cap.get(cv.CAP_PROP_FPS)
width = int(cap.get(cv.CAP_PROP_FRAME_WIDTH))
height = int(cap.get(cv.CAP_PROP_FRAME_HEIGHT))
font = cv.FONT_HERSHEY_SIMPLEX
pos = (10, height - 10)
print(width,height,fps)

firsttime=True

if not cap.isOpened():
    print("Cannot open camera")
    exit()
while True:
    # Capture frame-by-frame
    ret, frame = cap.read()
    # if frame is read correctly ret is True
    if not ret:
        print("Can't receive frame (stream end?). Exiting ...")
        break
    # Our operations on the frame come here
    gray = cv.cvtColor(frame, cv.COLOR_BGR2GRAY)
    gray = cv.GaussianBlur(gray,(3,3),0 )
    if cv.waitKey(1) == ord('q'):
        break
    if firsttime == False:      # we have a lastframe, lets compare
        diff = cv.absdiff(gray, lastgray)
        thresh = cv.threshold(diff, 25, 255, cv.THRESH_BINARY)[1]
        thresh = cv.dilate(thresh,None,iterations=2)
        cv.imshow('difference', thresh)
        cv.imshow('gray', gray)
        # Find the contours of the motion regions
        contours, hierarchy = cv.findContours(thresh, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE)
        # Draw bounding boxes around the contours
        totalsurface=0
        for c in contours:
            # Get the coordinates and dimensions of the bounding box
            x, y, w, h = cv.boundingRect(c)
            totalsurface += w*h
            # Draw a green rectangle on the original frame
            cv.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 1)
        totalsurface = totalsurface/width/height*100
        if 0 <= totalsurface <= 0.05:
            motion = "none"
        elif 0.05 <= totalsurface <= 0.5:
            motion = "slow"
        elif 0.5 <= totalsurface <= 1:
            motion = "medium"
        else:
            motion = "rapid"
        text= "change: "+ "{:02.03f}".format(totalsurface) + "% " + motion
        cv.putText(frame, text, pos, font, 1, (0, 255, 0), 2, cv.LINE_AA)
        cv.imshow("Motion Detection", frame)
    else:
        firsttime=False
        lastgray=gray
    lastgray=gray
    #dt = str(datetime.datetime.now().strftime("%H:%M:%S.%f"))
    # Put the timestamp on the frame
    #cv.putText(gray, dt, pos, font, 1, (255, 255, 255), 2, cv.LINE_AA)
    # Display the resulting frame
    
# When everything done, release the capture
cap.release()
cv.destroyAllWindows()