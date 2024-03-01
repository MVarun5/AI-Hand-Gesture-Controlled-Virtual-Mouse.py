# Importing packages

import cv2
import mediapipe as mp
import pyautogui
import math
from enum import IntEnum
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
from google.protobuf.json_format import MessageToDict
import screen_brightness_control as sbcontrol

pyautogui.FAILSAFE = False
mp_drawing = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands

# Assigning values for particular Gestures
class Gesture(IntEnum):
    fistm = 0
    pinkym = 1
    ringm = 2
    midm = 4
    last3m = 7
    indexm = 8
    first2m = 12
    last4m = 15
    thumbm = 16    
    palmm = 31
    
    Vgestm = 33
    closed2fing = 34
    pinchmajm = 35
    pinchminm = 36

# Multi-handedness Labeling
class MHlabel(IntEnum):
    minorv = 0
    majorv = 1

# Converting Mediapipe landmarks into recognizable hand gestures
class HandRecog:
    
    def __init__(self, hand_label):
        self.finger = 0
        self.ori_gesture = Gesture.palmm
        self.prev_gesture = Gesture.palmm
        self.frame_count = 0
        self.hand_result = None
        self.hand_label = hand_label
    
    def handResult(self, hand_result):
        self.hand_result = hand_result

    def numsignDist(self, point):
        numsign = -1
        if self.hand_result.landmark[point[0]].y < self.hand_result.landmark[point[1]].y:
            numsign = 1
        d = (self.hand_result.landmark[point[0]].x - self.hand_result.landmark[point[1]].x)**2
        d += (self.hand_result.landmark[point[0]].y - self.hand_result.landmark[point[1]].y)**2
        d = math.sqrt(d)
        return d*numsign
    
    def distance(self, point):
        d = (self.hand_result.landmark[point[0]].x - self.hand_result.landmark[point[1]].x)**2
        d += (self.hand_result.landmark[point[0]].y - self.hand_result.landmark[point[1]].y)**2
        d = math.sqrt(d)
        return d
    
    def dzDist(self,point):
        return abs(self.hand_result.landmark[point[0]].z - self.hand_result.landmark[point[1]].z)
    

    def fingstate(self):
        if self.hand_result == None:
            return

        points = [[8,5,0],[12,9,0],[16,13,0],[20,17,0]]
        self.finger = 0
        self.finger = self.finger | 0 #for thumb finger
        for idx,point in enumerate(points):
            
            d = self.numsignDist(point[:2])
            d2 = self.numsignDist(point[1:])
            
            try:
                ratio = round(d/d2,1)
            except:
                ratio = round(d1/0.01,1)

            self.finger = self.finger << 1
            if ratio > 0.5 :
                self.finger = self.finger | 1
    

    # Handling the Fluctations raised due to noise
    def getgest(self):
        if self.hand_result == None:
            return Gesture.palmm

        currgest = Gesture.palmm
        if self.finger in [Gesture.last3m,Gesture.last4m] and self.distance([8,4]) < 0.05:
            if self.hand_label == MHlabel.minorv :
                currgest = Gesture.pinchminm
            else:
                currgest = Gesture.pinchmajm

        elif Gesture.first2m == self.finger :
            point = [[8,12],[5,9]]
            d1 = self.distance(point[0])
            d2 = self.distance(point[1])
            ratio = d1/d2
            if ratio > 1.7:
                currgest = Gesture.Vgestm
            else:
                if self.dzDist([8,12]) < 0.1:
                    currgest =  Gesture.closed2fing
                else:
                    currgest =  Gesture.midm
            
        else:
            currgest =  self.finger
        
        if currgest == self.prev_gesture:
            self.frame_count += 1
        else:
            self.frame_count = 0

        self.prev_gesture = currgest

        if self.frame_count > 4 :
            self.ori_gesture = currgest
        return self.ori_gesture

# Executing commands according to the gestures
class Controller:
    tx_old = 0
    ty_old = 0
    trial = True
    flag = False
    grabflag = False
    pinch_major_flag = False
    pinch_minor_flag = False
    pinch_start_x = None
    pinch_start_y = None
    pinch_dir = None
    prev_pinch_lv = 0
    pinch_lv = 0
    framecount = 0
    prev_hand = None
    pinch_threshold = 0.3
    
    def pinch_y_lv(hand_result):
        d = round((Controller.pinch_start_y - hand_result.landmark[8].y)*10,1)
        return d

    def pinch_x_lv(hand_result):
        d = round((hand_result.landmark[8].x - Controller.pinch_start_x)*10,1)
        return d
    
    def sys_brightness():
        currbr_lv = sbcontrol.get_brightness()/100.0
        currbr_lv += Controller.pinch_lv/50.0
        if currbr_lv > 1.0:
            currbr_lv = 1.0
        elif currbr_lv < 0.0:
            currbr_lv = 0.0       
        sbcontrol.fade_brightness(int(100*currbr_lv) , start = sbcontrol.get_brightness())
    
    def sys_vol():
        devices = AudioUtilities.GetSpeakers()
        interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
        vol = cast(interface, POINTER(IAudioEndpointVolume))
        currvol_lv = vol.GetMasterVolumeLevelScalar()
        currvol_lv += Controller.pinch_lv/50.0
        if currvol_lv > 1.0:
            currvol_lv = 1.0
        elif currvol_lv < 0.0:
            currvol_lv = 0.0
        vol.SetMasterVolumeLevelScalar(currvol_lv, None)
    
    def scroll_Vertical():
        pyautogui.scroll(120 if Controller.pinch_lv>0.0 else -120)
        
    
    def scroll_Horizontal():
        pyautogui.keyDown('shift')
        pyautogui.keyDown('ctrl')
        pyautogui.scroll(-120 if Controller.pinch_lv>0.0 else 120)
        pyautogui.keyUp('ctrl')
        pyautogui.keyUp('shift')

    # Showing hand to get the cursor position
    def get_pos(hand_result):
        point = 9
        pos = [hand_result.landmark[point].x ,hand_result.landmark[point].y]
        sx,sy = pyautogui.size()
        x_old,y_old = pyautogui.position()
        x = int(pos[0]*sx)
        y = int(pos[1]*sy)
        if Controller.prev_hand is None:
            Controller.prev_hand = x,y
        delta_x = x - Controller.prev_hand[0]
        delta_y = y - Controller.prev_hand[1]

        d_sq = delta_x**2 + delta_y**2
        ratio = 1
        Controller.prev_hand = [x,y]

        if d_sq <= 25:
            ratio = 0
        elif d_sq <= 900:
            ratio = 0.07 * (d_sq ** (1/2))
        else:
            ratio = 2.1
        x , y = x_old + delta_x*ratio , y_old + delta_y*ratio
        return (x,y)

    def pinch_control_init(hand_result):
        Controller.pinch_start_x = hand_result.landmark[8].x
        Controller.pinch_start_y = hand_result.landmark[8].y
        Controller.pinch_lv = 0
        Controller.prev_pinch_lv = 0
        Controller.framecount = 0

    # Hold final position for 5 frames to change status
    def pinch_control(hand_result, controlHorizontal, controlVertical):
        if Controller.framecount == 5:
            Controller.framecount = 0
            Controller.pinch_lv = Controller.prev_pinch_lv

            if Controller.pinch_dir == True:
                controlHorizontal() # for x value

            elif Controller.pinch_dir == False:
                controlVertical() #for y value

        lvx =  Controller.pinch_x_lv(hand_result)
        lvy =  Controller.pinch_y_lv(hand_result)
            
        if abs(lvy) > abs(lvx) and abs(lvy) > Controller.pinch_threshold:
            Controller.pinch_dir = False
            if abs(Controller.prev_pinch_lv - lvy) < Controller.pinch_threshold:
                Controller.framecount += 1
            else:
                Controller.prev_pinch_lv = lvy
                Controller.framecount = 0

        elif abs(lvx) > Controller.pinch_threshold:
            Controller.pinch_dir = True
            if abs(Controller.prev_pinch_lv - lvx) < Controller.pinch_threshold:
                Controller.framecount += 1
            else:
                Controller.prev_pinch_lv = lvx
                Controller.framecount = 0

    def handle_controls(gest, hand_result):        
        x,y = None,None
        if gest != Gesture.palmm :
            x,y = Controller.get_pos(hand_result)
        
        # flag reset
        if gest != Gesture.fistm and Controller.grabflag:
            Controller.grabflag = False
            pyautogui.mouseUp(button = "left")

        if gest != Gesture.pinchmajm and Controller.pinch_major_flag:
            Controller.pinch_major_flag = False

        if gest != Gesture.pinchminm and Controller.pinch_minor_flag:
            Controller.pinch_minor_flag = False

        # implementation
        if gest == Gesture.Vgestm:
            Controller.flag = True
            pyautogui.moveTo(x, y, duration = 0.1)

        elif gest == Gesture.fistm:
            if not Controller.grabflag : 
                Controller.grabflag = True
                pyautogui.mouseDown(button = "left")
            pyautogui.moveTo(x, y, duration = 0.1)

        elif gest == Gesture.midm and Controller.flag:
            pyautogui.click()
            Controller.flag = False

        elif gest == Gesture.indexm and Controller.flag:
            pyautogui.click(button='right')
            Controller.flag = False

        elif gest == Gesture.closed2fing and Controller.flag:
            pyautogui.doubleClick()
            Controller.flag = False

        elif gest == Gesture.pinchminm:
            if Controller.pinch_minor_flag == False:
                Controller.pinch_control_init(hand_result)
                Controller.pinch_minor_flag = True
            Controller.pinch_control(hand_result,Controller.scroll_Horizontal, Controller.scroll_Vertical)
        
        elif gest == Gesture.pinchmajm:
            if Controller.pinch_major_flag == False:
                Controller.pinch_control_init(hand_result)
                Controller.pinch_major_flag = True
            Controller.pinch_control(hand_result,Controller.sys_brightness, Controller.sys_vol)
        
'''
----------------------------------------  Main Class  ----------------------------------------

'''


class GestureController:
    gc_mode = 0
    cap = None
    cam_h = None
    cam_w = None
    hr_major = None # It takes Right Hand by default
    hr_minor = None # It takes Left hand by default
    dom_hand = True

    def __init__(self):
        GestureController.gc_mode = 1
        GestureController.cap = cv2.VideoCapture(0)
        GestureController.cam_h = GestureController.cap.get(cv2.CAP_PROP_FRAME_HEIGHT)
        GestureController.cam_w = GestureController.cap.get(cv2.CAP_PROP_FRAME_WIDTH)
    
    def classify_hands(results):
        l , r = None,None
        try:
            handedness_dict = MessageToDict(results.multi_handedness[0])
            if handedness_dict['classification'][0]['label'] == 'Right':
                r = results.multi_hand_landmarks[0]
            else :
                l = results.multi_hand_landmarks[0]
        except:
            pass

        try:
            handedness_dict = MessageToDict(results.multi_handedness[1])
            if handedness_dict['classification'][0]['label'] == 'Right':
                r = results.multi_hand_landmarks[1]
            else :
                l = results.multi_hand_landmarks[1]
        except:
            pass
        
        if GestureController.dom_hand == True:
            GestureController.hr_major = r
            GestureController.hr_minor = l
        else :
            GestureController.hr_major = l
            GestureController.hr_minor = r

    def start(self):
        
        handmajor = HandRecog(MHlabel.majorv)
        handminor = HandRecog(MHlabel.minorv)

        with mp_hands.Hands(max_num_hands = 2,min_detection_confidence=0.5, min_tracking_confidence=0.5) as hands:
            while GestureController.cap.isOpened() and GestureController.gc_mode:
                success, image = GestureController.cap.read()

                if not success:
                    print("Ignoring empty camera frame.")
                    continue
                
                image = cv2.cvtColor(cv2.flip(image, 1), cv2.COLOR_BGR2RGB)
                image.flags.writeable = False
                results = hands.process(image)
                
                image.flags.writeable = True
                image = cv2.cvtColor(image, cv2.COLOR_RGB2BGR)

                if results.multi_hand_landmarks:                   
                    GestureController.classify_hands(results)
                    handmajor.handResult(GestureController.hr_major)
                    handminor.handResult(GestureController.hr_minor)

                    handmajor.fingstate()
                    handminor.fingstate()
                    gest_name = handminor.getgest()

                    if gest_name == Gesture.pinchminm:
                        Controller.handle_controls(gest_name, handminor.hand_result)
                    else:
                        gest_name = handmajor.getgest()
                        Controller.handle_controls(gest_name, handmajor.hand_result)
                    
                    for hand_landmarks in results.multi_hand_landmarks:
                        mp_drawing.draw_landmarks(image, hand_landmarks, mp_hands.HAND_CONNECTIONS)
                else:
                    Controller.prev_hand = None
                cv2.imshow('Gesture Controller', image)
                if cv2.waitKey(5) & 0xFF == 13:
                    break
        GestureController.cap.release()
        cv2.destroyAllWindows()

gc1 = GestureController()
gc1.start()
