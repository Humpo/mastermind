exceptions={#this list holds the ascii values of important keyboard values
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
def clickL(x,y):#left clicks based on x and y coordinates on screen
    win32api.SetCursorPos((x,y))#sets cursor position
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)#
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)#clicks down then up
def clickR(x,y):#right clicks based on x and y coordinates on screen
    win32api.SetCursorPos((x,y))
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0)#
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0)#clicks down then up
def Main():
    filename=raw_input("filename: ")
    try:filestuff=open(filename+'.txt','r')
    except FileNotFoundError:print "File: "+filename+", Not found.\n Exiting..."
    print "Running: "+' '+filename
    #initialize lists of basic key inputs
    chars=[]#store keyboard characters
    times_char=[]#Time between instances
    times_clickL=[]#Timed clicked left
    times_clickR=[]#'' '' right
    xandyL=[]#x and y location for left click
    xandyR=[]#'' '' right click
    for line in filestuff:#read each line in specially created text file
        char=None#initialize the possible charater to be initiated
        temp = re.split('\s',line)#seperated inputs into a list based on space syntax in text file
        if temp[0] == 'clickL':#clicks left
            #line ['clickL', X location, Y Location, time]
            times_clickL.append(float(temp[3]))#adds time to time queue
            xandyL.append([int(temp[1]),int(temp[2])])
            #queue.enter(delay=time_start,0,action=click,arg=(x_loc,y_loc))
        elif temp[0] == 'clickR':#clicks right
            #list contents same at 'clickL'
            x_loc=temp[1]
            y_loc=temp[2]
            times_clickR.append(float(temp[3])) #time (ms)
            xandyR.append([int(x_loc),int(y_loc)])
            #queue.enter(delay=time_start,0,action=click,arg=(x_loc,y_loc))
        else:
            #this is a small exception list in order to fix small bugs due to the amount of time working during the hackathon
            if temp[0] == "enter":char='\r'
            elif temp[0] == "tab":char ='\t'
            elif temp[0] == "space":char = ' '
            elif len(temp[0]) == 1:char = temp[0]
            else:print("ERROR, bad file?")#incase bad things happen
            #special cases are shown as string indicates
            times_char.append(float(temp[1]))#adds the time to time list
            chars.append(char)#adds the character to the char list
    s = sched.scheduler(time.time,time.sleep)#create a scheduler instance which will run events at recorded time
    for i in range(len(times_char)):
        #print times_char[i]
        #print charsL[i]
        s.enter(times_char[i],1,outkey,(chars[i], exceptions))
    for i in range(len(times_clickL)):
        #print times_clickL[i]
        #print xandyL[i]
        s.enter(times_clickL[i],1,clickL,(xandyL[i][0], xandyL[i][1]))
    for i in range(len(times_clickR)):
        #print times_clickR[i]
        #print xandyR[i]
        s.enter(times_clickR[i],1,clickR,(xandyR[i][0], xandyR[i][1]))
    #queue data into 's' scheduler instance
    s.run()#run the whole schedule
if __name__=="__main__":
    import win32com.client 
    from win32api import keybd_event
    import win32api,win32con
    from threading import Timer
    import time
    import sched
    import re
    Main()
    #created at HACKRPI 2014 by Alexander Comerford 
