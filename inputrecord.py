def onclickL(event):
    end = time.time()
    (mx, my) = event.Position
    rect = win32gui.GetWindowRect(win32gui.GetForegroundWindow())
    x = rect[0]
    y = rect[1]
    w = rect[2] - x
    h = rect[3] - y
    f.write("clickL " + str(mx) + " " + str (my) + " ")
    f.write(str(end - start) + "\n")
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
    f.write(str(end - start) + "\n")
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
        elif char == "~":         
            f.close()
            from sys import exit
            exit(0)
        else:
            f.write(char)
        f.write(" " + str(end - start) + "\n")
    return True

def Main():
    name=raw_input('Macroname: ') + ".txt"
    global start, f
    start = time.time()
    f = open(name,"w")#creates new file
    event = pyHook.HookManager()      #open pyHook
    event.KeyDown = pressed_chars     #set pressed_chars function on KeyDown event
    event.HookKeyboard()              #start the function
    event.MouseLeftDown=onclickL      #onclickL function generated on left mouse
    event.MouseRightDown=onclickR     #onclickR function generated on Right mouse
    event.HookMouse()                 #queue events
    pythoncom.PumpMessages()          #"pump messages" or run the queue
    event.UnhookMouse()               #kills the mouse record
if __name__=="__main__":
    import time
    import win32gui
    import win32con
    import win32ui
    import pyHook
    import pythoncom
    Main()
    #created at HACKRPI 2014 by Alexander Comerford 
