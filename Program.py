import sys

if sys.platform in ['Windows', 'win32', 'cygwin']:
    import cv2 
    import time
    import pyautogui
    from pynput import keyboard
    from ctypes import *
    import json
    import win32gui 
    from pynput.mouse import Listener
    import os
    import datetime
    from threading import Thread
    import shutil
    import tkinter

    from dotenv import load_dotenv
    load_dotenv('.env')
    
    if not os.path.exists(os.getcwd()+"\\.env"):
        f= open('.env','w')
        f.write('TIME_INTERVAL = 10\n')
        f.write('PATH_DIR='+os.getcwd())
        f.close()
    

    TIME_INTERVAL=int(os.getenv('TIME_INTERVAL'))  #in seconds
    DIR_PATH= str(os.getenv('PATH_DIR'))

    """ Directory """

    CURR_DATE= datetime.datetime.now().strftime('%d-%m-%Y')
    PATH = DIR_PATH+'\\Logs\\' + CURR_DATE

    if not os.path.exists(PATH):
        os.makedirs(PATH)
        os.makedirs( PATH +'\\SS')
        os.makedirs( PATH +'\\WC')
        os.makedirs( PATH +'\\Raw-data')
        os.makedirs( PATH + '\\Aggregated-data')
    """ Directory End """

    """ For Mouse Clicks """

    global pre_window_mouse
    global counter
    global start_time
    global previous_pid
    previous_pid=0
    global data
    global data1
    global t_counter
    data = []
    data1 = [] #switch
    t_counter = 0

    def on_click(x, y, button, pressed):
        if pressed:
            global pre_window_mouse
            global counter
            global start_time
            global previous_pid
            global data,data1,t_counter
            pid = c_ulong(0)
            window = win32gui.GetForegroundWindow()
            windll.user32.GetWindowThreadProcessId(windll.user32.GetForegroundWindow(),byref(pid))
            # While beginnig for the first time
            if previous_pid == 0:
                pre_window_mouse=window
                counter=0
                start_time = datetime.datetime.now().strftime('%d/%m/%Y - %H:%M:%S')
                previous_pid = pid.value
            # When application is switched
            if pre_window_mouse!=window:
                data1.append({'PID':previous_pid,'Application':win32gui.GetWindowText(pre_window_mouse),'StartTime':start_time,'EndTime':datetime.datetime.now().strftime('%H:%M:%S'),'ClickCounter':counter})
                pre_window_mouse=window
                counter=0
                start_time = datetime.datetime.now().strftime('%d/%m/%Y - %H:%M:%S')
                previous_pid= windll.user32.GetWindowThreadProcessId(windll.user32.GetForegroundWindow(),byref(pid))
            # When click event occurs
            data.append({'PID':str(pid.value),'Application':str(win32gui.GetWindowText(window)),'TimeStamp':datetime.datetime.now().strftime('%H:%M:%S'),'MouseX':x,'MouseY':y})
            counter+=1
            t_counter+=1
            if t_counter%5 == 0:
                write()
                data = []
                data1 = []
                t_counter=0
            release_click= False

    def write():
        global data,data1
        with open(PATH +"\\Raw-data\\Mouse_Raw_Logs.txt",'a+') as f:
            for d in data:
                json.dump(d,f)
                f.write('\n')
        with open(PATH +"\\Aggregated-data\\Mouse_Logs.txt",'a+') as f1:
            for d1 in data1:
                json.dump(d1,f1)
                f1.write('\n')

    def mouse():
        with Listener(on_click=on_click) as listener:
            listener.join()

    """ For Mouse Clicks """

    """ For Web Cam """

    def Webcam():
        i = 0
        while(1):
            cap = cv2.VideoCapture(0)
            ret, frame = cap.read()
            
            cv2.imwrite(PATH+'\\WC\\IMG_'+datetime.datetime.now().strftime('%d-%m-%Y_%H%M%S')+'.png', frame)
            cap.release()
            time.sleep(TIME_INTERVAL)

    """ For Web Cam """

    """ For Screenshots """

    def Screenshot():
        while(1):
            pic = pyautogui.screenshot()
            pic.save(PATH+'\\SS\\Screenshot_'+datetime.datetime.now().strftime('%d-%m-%Y_%H%M%S')+'.png')
            #pyautogui.PAUSE = 1000
            time.sleep(TIME_INTERVAL)

    """ For Screenshots """

    """ For Keyboard Clicks """
    pre_window= None
    active_window= None
    c=0
    data_list=[]
    raw_data_dict_list=[]
    data_dict_list=[]
    t_count=0

    def on_press(key):
        global c, pre_window, active_window, data_list,t_count, raw_data_dict_list, data_dict_list
        c+=1
        t_count+=1

        if pre_window == None:
             pre_window=win32gui.GetForegroundWindow()
             active_window = pre_window
             
        else:
            pre_window = active_window
            active_window = win32gui.GetForegroundWindow()
            
        #print(active_window,key,sep='\t')
        window_name=win32gui.GetWindowText(pre_window)
        pid = c_ulong(0)
        windll.user32.GetWindowThreadProcessId(windll.user32.GetForegroundWindow(),byref(pid))        
        
        try:
            character='{0}'.format(key.char)
        except Exception as e:
            character=' Functional key '

        raw_data_dict_list.append({'PID': str(pid.value),'Application': window_name, 'Timestamp': datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'),'Key Pressed: ': character, })

        if pre_window!=active_window :
            
            try:
                
                #print(data_list)
                #data_dict={'PID': str(pid.value),'Application': window_name, 'Timestamp': datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'),'Key Pressed: ':''.join( data_list), 'Keyboard Clicks': str(c) }
                

                data_dict_list.append({'PID': str(pid.value),'Application': window_name, 'Timestamp': datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S'),'Key Pressed: ':''.join( data_list), 'Keyboard Clicks': str(c) })
                c=0
                
                data_list=[]
                    #f.write('Key Pressed: '+''.join( data_list)+'\nApplication: '+active_window+'\nTimeStamp: '+datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')+'\nPID: '+str(pid.value)+'\n\n')
            except Exception as e:
                print(e)
                    
        else:
            data_list.append(character)
            
        if t_count%5==0:
            t_count=0
            #global data_dict_list
            #globalaaaa raw_data_dict_list

            with open(PATH+'\\Aggregated-data\\'+'Keyboard_Logs.txt','a+') as f:
                for d in data_dict_list:
                    
                    try:
                        json.dump(d,f)
                    except Exception as e:
                        print(e)
                    f.write('\n')
            data_dict_list=[]

            with open(PATH+'\\Raw-data\\'+'Keyboard_Raw-Logs.txt','a+') as f:
                for d in raw_data_dict_list:
                    try:
                        json.dump(d,f)
                    except Exception as e:
                        print(e)
                        print(d)
                    f.write('\n')
            raw_data_dict_list=[]

            
    def KeyBoardHit():
        with keyboard.Listener(on_press=on_press) as listenerk:
            listenerk.join()
            
    """ For Keyboard Clicks """

    def Start():
        Thread(target = Screenshot).start()
        Thread(target = Webcam).start()
        Thread(target = mouse).start()
        Thread(target = KeyBoardHit).start()
        return

    def Stop():
        top.destroy()
        os.system('Taskkill /PID '+str(os.getpid())+' /F')

    if __name__ == '__main__':
        try:
            global top
            top = tkinter.Tk()
            B = tkinter.Button(top, text ="Start", command = Start).pack()
            B1 = tkinter.Button(top, text ="Stop", command = Stop).pack()
            top.mainloop()
        except Exception as e:
            print(e)
        
        #B1 = tkinter.Button(top, text ="Pause", command = Pause).pack()
elif sys.platform in ['linux', 'linux2']:
    print('You are using Linux. Please run it on Windows')
elif sys.platform in ['Mac', 'darwin', 'os2', 'os2emx']:
    print('You are using MacOS. Please run it on Windows')
