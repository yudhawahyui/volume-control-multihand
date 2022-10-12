import cv2 as cv
import numpy as np
from cvzone.HandTrackingModule import HandDetector
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(
    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
volume = cast(interface, POINTER(IAudioEndpointVolume))

# volume.GetMute()
# volume.GetMasterVolumeLevel()
volRange = volume.GetVolumeRange()
# volume.SetMasterVolumeLevel(-20.0, None)

minVol = volRange[0]
maxVol = volRange[1]
vol = 0
vol_bar = 400
volPer = 0

cap = cv.VideoCapture(0)
detector = HandDetector(detectionCon=0.8, maxHands=2)
while True:
    success, img = cap.read()
    img = cv.flip(img, 1)
    hands,img = detector.findHands(img, flipType=False)

    if len(hands) == 1: 
        # Hand 1
        hand1 = hands[0]
        lmList1 = hand1["lmList"] # List of 21 Landmark points
        bbox1 = hand1["bbox"] # Bounding box info x,y,w,h
        centerPoint1 = hand1['center'] # center of the hand cx,cy
        handType1 = hand1["type"] # Handtype Left or Right

        fingers1 = detector.fingersUp(hand1)
        length, info, img = detector.findDistance(lmList1[8][0:2], lmList1[4][0:2], img)
        
        # Kotak Volume
        vol = np.interp(length, [50, 140], [minVol,maxVol])
        vol_bar = np.interp(length, [50, 140], [400,150])
        volPer = np.interp(length, [50, 140], [0,100])
        # print(int(length),vol_bar)
        volume.SetMasterVolumeLevel(vol, None)

        cv.rectangle(img, (50, int(vol_bar)), (85, 400), (0, 255, 0), cv.FILLED)
        cv.rectangle(img, (50, 145), (85, 400), (0, 0, 0), 3)
        cv.putText(img, f'{int(volPer)}%' , (40, 450), cv.FONT_HERSHEY_PLAIN, 2, (0, 250, 0), 3)
    
    elif len(hands) == 2:
        # Hand 2
        hand1 = hands[0]
        hand2 = hands[1]
        lmList1 = hand1["lmList"]
        lmList2 = hand2["lmList"]
        bbox2 = hand2["bbox"]
        centerPoint2 = hand2['center']
        handType2 = hand2["type"]

        fingers2 = detector.fingersUp(hand2)
        length, info, img = detector.findDistance(lmList1[8][0:2], lmList2[8][0:2], img)
        
        # Kotak Volume
        vol = np.interp(length, [50, 140], [minVol,maxVol])
        vol_bar = np.interp(length, [50, 140], [400,150])
        volPer = np.interp(length, [50, 140], [0,100])
        # print(int(length),vol_bar)
        volume.SetMasterVolumeLevel(vol, None)

        cv.rectangle(img, (50, int(vol_bar)), (85, 400), (0, 255, 0), cv.FILLED)
        cv.rectangle(img, (50, 145), (85, 400), (0, 0, 0), 3)
        cv.putText(img, f'{int(volPer)}%' , (40, 450), cv.FONT_HERSHEY_PLAIN, 2, (0, 250, 0), 3)

    cv.imshow('Hand Detector', img)
    if cv.waitKey(1) == ord('q'):
        break

cap.release()
cv.destroyAllWindows()

