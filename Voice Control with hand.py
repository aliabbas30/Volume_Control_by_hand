import cv2
import mediapipe as mp
import time
import math
import numpy as np
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
########################
wCam, hCam = 640, 480
########################

devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = interface.QueryInterface(IAudioEndpointVolume)
# volume.GetMute()
# volume.GetMasterVolumeLevel()
volume_range = volume.GetVolumeRange()
# volume.SetMasterVolumeLevel(0.0, None)
minVol = volume_range[0]
maxVol = volume_range[1]
print(minVol, maxVol)


cap = cv2.VideoCapture(0)
cap.set(3, wCam)
cap.set(4, hCam)

mpHands = mp.solutions.hands
# ONLY USE RGB
hands = mpHands.Hands()
mp_draw = mp.solutions.drawing_utils

pTime = 0
cTime = 0
volBar = 400
volpercentage = 0
while True:
    success, img = cap.read()
    img_rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
    results = hands.process(img_rgb)
    #     print(results.multi_hand_landmarks)
    ## single hand detection
    if results.multi_hand_landmarks:
        lmlist = []
        for handlms in results.multi_hand_landmarks:
            for ids, lm in enumerate(handlms.landmark):
                ## get the pixel values
                h, w, c = img.shape
                cx, cy = int(lm.x * w), int(lm.y * h)
                lmlist.append([ids, cx, cy])
                cv2.circle(img, (cx, cy), 5, (255, 0, 24), cv2.FILLED)
                mp_draw.draw_landmarks(img, handlms, mpHands.HAND_CONNECTIONS)
        if len(lmlist) != 0:
            x1, y1 = lmlist[4][1], lmlist[4][2]
            x2, y2 = lmlist[8][1], lmlist[8][2]
            cx,cy =(x1+x2)//2, (y1+y2)//2
            cv2.circle(img, (x1, y1), 15, (255,0,0),cv2.FILLED)
            cv2.circle(img, (x2, y2), 15, (255, 0, 0), cv2.FILLED)
            cv2.line(img, (x1,y1) ,(x2, y2) ,(255,0,255), 2)
            cv2.circle(img, (cx, cy), 15, (255, 0, 0), cv2.FILLED)
            length = math.hypot(x2-x1, y2 - y1)
            # hand range 20-150
            # volume range = -74 to 0
            vol = np.interp(length, [20,150], [-74,0])
            volBar = np.interp(length, [20, 150], [400, 150])
            volpercentage = np.interp(length, [20, 150], [0, 100])
            volume.SetMasterVolumeLevel(vol, None)
            if length < 40:
                cv2.circle(img, (cx, cy), 15, (255, 255, 255), cv2.FILLED)
    cv2.rectangle(img,(50,150),(85,400),(0,255,0),3)
    cv2.rectangle(img,(50,int(volBar)),(85,400),(255,0,0),cv2.FILLED)
    cv2.putText(img, f'Volume Percentage: {int(volpercentage)} %', (40, 450), cv2.FONT_HERSHEY_SIMPLEX, 1,
                (255, 0, 0), 2)

    cTime = time.time()
    fps = 1 / (cTime - pTime)
    pTime = cTime
    cv2.putText(img, f'FPS: {int(fps)}', (40, 50), cv2.FONT_HERSHEY_COMPLEX, 1,
                (255, 0, 0), 3)
    cv2.imshow('Image', img)
    cv2.waitKey(1)
