#py2

import win32api, time, win32gui, win32con, win32ui, pyHook, pythoncom
name=raw_input('Macroname: ')
log_file = name + ".txt"               

start = time.time()
f = open(log_file,"w") 

def onclickL(event):
    end = time.time()
    (mx, my) = event.Position
    rect = win32gui.GetWindowRect(win32gui.GetForegroundWindow())
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    f.write("clickL " + str(mx) + " " + str (my) + " ")
    f.write(str(end - start) + " ")
    f.write("P[" + str(x) + " " + str(y) + "] ")
    f.write("S[" + str(w) + " " + str(h)+ "] ")
    f.write("W[" + win32gui.GetWindowText(win32gui.GetForegroundWindow()) + "]\n")
    return True

def onclickR(event):
    end = time.time()
    (mx, my) = event.Position
    rect = win32gui.GetWindowRect(win32gui.GetForegroundWindow())
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    f.write("clickR " + str(mx) + " " + str (my) + " ")
    f.write(str(end - start) + " ")
    f.write("P[" + str(x) + " " + str(y) + "] ")
    f.write("S[" + str(w) + " " + str(h)+ "] ")
    f.write("W[" + win32gui.GetWindowText(win32gui.GetForegroundWindow()) + "]\n")
    return True


def pressed_chars(event):
    end = time.time()
    if event.Ascii:
        char = chr(event.Ascii) 
        if event.Ascii == 13:
            f.write("enter")
        elif event.Ascii == 9:
            f.write("tab")
        elif event.Ascii == 32:
            f.write("space")
        elif char == "|":         
            f.close()          
            exit()              
        else:
            f.write(char)
        f.write(" " + str(end - start) + "\n")
    return True



proc = pyHook.HookManager()      #open pyHook
proc.KeyDown = pressed_chars     #set pressed_chars function on KeyDown event
proc.HookKeyboard()              #start the function
proc.MouseLeftDown=onclickL
proc.MouseRightDown=onclickR
proc.HookMouse()
pythoncom.PumpMessages()
proc.UnhookMouse()
