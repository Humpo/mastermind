import win32com.client 
from win32api import keybd_event
import win32api,win32con
from threading import Timer
import time
import sched
import re
import ctypes
exceptions={
    'END': 35,
    'DOWN': 40,
    'LEFT': 37,
    'UP': 38,
    'RIGHT': 39,
    'RCTRL': 163,
    'CTRL': 17,
    'ESC': 27,
    ' ': 32,
    'VOLUP': 175,
    'DOLDOWN': 174,
    'NUMLOCK': 144,
    'SCROLL': 145,
    'SELECT': 41,
    'PRINTSCR': 44,
    'INS': 45,
    'DEL': 46,
    'LWIN': 91,
    'RWIN': 92,
    'ALT': 18,
    'TAB': 9,
    'CAPSLOCK': 20,
    'ENTER': 13,
    'BS': 8,
    'LSHIFT': 160,
    'SHIFT': 161,
    'LCTRL': 162,
    }
def outkey(key, exceptions):
    try:
        win32com.client.Dispatch("WScript.Shell").SendKeys(key)
    except:
        if key in exceptions:
            win32com.client.Dispatch("WScript.Shell").SendKeys(chr(exceptions[key]))
def clickL(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
    #ctypes.windll.user32.SetCursorPos(x, y)
    #ctypes.windll.user32.mouse_event(2, 0, 0, 0,0) # left down
    #ctypes.windll.user32.mouse_event(4, 0, 0, 0,0) # left up
def clickR(x,y):
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0)
filename=raw_input("filename: ")
print filename
filestuff=open(filename+'.txt','r')
chars=[]
times_char=[]
times_clickL=[]
times_clickR=[]
xandyL=[]
xandyR=[]
for line in filestuff:
    char=None
    temp = re.split('\s',line)
    if temp[0] == 'clickL':
        x_loc=temp[1]
        y_loc=temp[2]
        times_clickL.append(float(temp[3])) #time (ms)
        xandyL.append([int(x_loc),int(y_loc)])
        #queue.enter(delay=time_start,0,action=click,arg=(x_loc,y_loc))
    elif temp[0] == 'clickR':
        x_loc=temp[1]
        y_loc=temp[2]
        times_clickR.append(float(temp[3])) #time (ms)
        xandyR.append([int(x_loc),int(y_loc)])
        #queue.enter(delay=time_start,0,action=click,arg=(x_loc,y_loc))
    else:
        if temp[0] == "enter":
            char='\r'
        elif temp[0] == "tab":
            char ='\t'
        elif temp[0] == "space":
            char = ' '
        elif len(temp[0]) == 1:
            char = temp[0]
        else:
            print("ERROR, bad file?")
        times_char.append(float(temp[1])) #time (ms)
        chars.append(char)
        
        #queue.enter(delay=time_start,0,action=outkey,arg=(char))
        
s = sched.scheduler(time.time,time.sleep)
for i in range(len(times_char)):
    print times_char[i]
    print charsL[i]
    s.enter(times_char[i],1,outkey,(chars[i], exceptions))
    #Timer(times_char[i],outkey,args=[chars[i]]).start()
for i in range(len(times_clickL)):
    print times_clickL[i]
    print xandyL[i]
    s.enter(times_clickL[i],1,clickL,(xandyL[i][0], xandyL[i][1]))
    #Timer(times_click[i],click,args=[xandy[i][0], xandy[i][1]]).start()
for i in range(len(times_clickR)):
    print times_clickR[i]
    print xandyR[i]
    s.enter(times_clickR[i],1,clickR,(xandyR[i][0], xandyR[i][1]))
    #Timer(times_click[i],click,args=[xandy[i][0], xandy[i][1]]).start()
s.run()
