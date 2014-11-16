#from _future_ import print_function
import httplib
import httplib2
import argparse
import webbrowser
from apiclient.discovery import build
import logging
logging.basicConfig()
from datetime import datetime, timedelta
import re
from threading import Timer
import time
import os

from oauth2client import tools
parser = argparse.ArgumentParser(parents=[tools.argparser])
flags = parser.parse_args()

#event setup
#Event_type:Event_command:Config (optional)
event_types = ["MACRO","WEBPAGE"]

#inputs
cal = raw_input("Calender: (google, office, or none) ")
deltime = timedelta(minutes=60)

if os.name == 'nt':
	#Alex's Code
	import win32com.client 
	from win32api import keybd_event
	import win32api,win32con
	import pythoncom
	import sched

	def outkey(key):
		pythoncom.CoInitialize()
		win32com.client.Dispatch("WScript.Shell").SendKeys(key)
	def clickL(x,y):
		pythoncom.CoInitialize()
		win32api.SetCursorPos((x,y))
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
		#ctypes.windll.user32.SetCursorPos(x, y)
		#ctypes.windll.user32.mouse_event(2, 0, 0, 0,0) # left down
		#ctypes.windll.user32.mouse_event(4, 0, 0, 0,0) # left up
	def clickR(x,y):
		pythoncom.CoInitialize()
		win32api.SetCursorPos((x,y))
		win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,x,y,0,0)
		win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,x,y,0,0)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
	def run_macro(event):
		pythoncom.CoInitialize()
		filestuff=open(event,'r')
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
				
		s = sched.scheduler(time.time,time.sleep)
		for i in range(len(times_char)):
			print times_char[i]
			print chars[i]
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
			#Timer(times_click[i],click,args=[xandy[i][0], xandy[i][1]]).start())
		s.run()
	#End Alex's Code
	
#Other commands
def run_webpage(url, etime):
	print "Opening: ", url, " at time: ", etime
	webbrowser.open(url)

if(cal=="google"):
	from oauth2client.client import OAuth2WebServerFlow
	from oauth2client.file import Storage
	from oauth2client.tools import run_flow
elif(cal=="office"):
	import requests
	import getpass

def main():
	if(cal=="google"):
		calID = raw_input("Enter google calendar ID: ")
		o2storage = Storage("./credentials.dat")
		credentials = o2storage.get()
		if credentials is None or credentials.invalid:
			flow = OAuth2WebServerFlow(client_id="304557316781-fk7lgpitio9mc536krgdo2bi8u2lickn.apps.googleusercontent.com",
			                   client_secret="nv1505W8edUll_k0qNAwKWaz",
			                   scope="https://www.googleapis.com/auth/calendar",
			                   redirect_uri="urn:ietf:wg:oauth:2.0:oob")
			credentials = run_flow (flow, o2storage,flags)

		ntime = datetime.now()
		minTime = ntime 
		maxTime = ntime + timedelta(hours=24)
		print time.strftime ("%Y-%m-%dT%H:%M:%S")
		minT = minTime.strftime ("%Y-%m-%dT%H:%M:%S")
		maxT = maxTime.strftime ("%Y-%m-%dT%H:%M:%S")

		minT += "-05:00" #we assume you live in New York #YOLO
		maxT += "-05:00"

		http = httplib2.Http()
		http = credentials.authorize(http)	
		service = build('calendar', 'v3', http=http)
	
		raw = str(service.events().list(calendarId=calID,timeMin=minT, timeMax=maxT).execute())

		temp = re.findall("u\'summary\': u\'[\S]*",raw) #counting number of events (+1)
		index = []
		command_queue = []
		event_start = []
		for i in range(len(temp)):
			event_name = re.split(';',re.split('\'',temp[i])[3])
			print "event canidate: ", event_name
			if event_name[0] in event_types: #compares event summary to the desired name of the event
				index.append(i) #counting number of matching events
				command_queue.append(event_name)
		temp = re.findall("u'start': {u\'dateTime\': u\'[\S]*",raw)
		for i in range(len(index)):
			event_start.append(re.split('\'',temp[index[i]])[5]) #grabs the times of matching
	
		eventTime = datetime.strptime(event_start[0], "%Y-%m-%dT%H:%M:%S-05:00")
		schedTime = eventTime - ntime
		for i in range(len(index)):
			print "event match: ", command_queue[i]
			eventTime = datetime.strptime(event_start[i], "%Y-%m-%dT%H:%M:%S-05:00")
			tmpTime = eventTime - ntime
			print "event time: ", tmpTime
			if tmpTime < schedTime:
				schedTime = tmpTime
	
	elif(cal=="office"):
		EMAIL = raw_input("Username: ")
		PASSWORD = getpass.getpass("Password: ")
		response = requests.get('https://outlook.office365.com/EWS/OData/users/'+EMAIL+'/Events', auth=(EMAIL, PASSWORD))
		ntime = datetime.now()
		x = 0
		i = 0
		command_queue = []
		for event in response.json()['value']:
  			print(event['Subject'])
  			command_queue.append(re.split(';',event['Subject']))
  			if command_queue[i][0] in event_types:
  				print "event match: ", command_queue[i]
  				x+= 1
  			i += 1
  			eventTime = datetime.strptime(event['Start'], "%Y-%m-%dT%H:%M:%SZ") - timedelta(hours=5) 
  			print eventTime
  			if (eventTime-ntime) < timedelta(hours=1) and (eventTime-ntime) > timedelta(hours=0):
  				i = x
  		
  	else: #this skips the calendar and just runs a command
  		raw = raw_input("Input cal string: ")
  		i = 0
  		command_queue = []
  		command_queue.append(re.split(';',raw))
  		ntime = datetime.now()
  		eventTime = ntime + timedelta(seconds=5) #5 second countdown
  		
	schedTime = eventTime - ntime
	if (schedTime > timedelta(hours=0)):
		print "Time until next event: ", schedTime
	else: 
		print "No events detected today, sleeping for 24hrs"
		return(timedelta(hours=24))
	if command_queue[i][0] == "MACRO":
		Timer(schedTime.total_seconds(), run_macro,args=[command_queue[i][1]] ).start()
	elif command_queue[i][0] == "WEBPAGE":
		Timer(schedTime.total_seconds(), run_webpage,args=[command_queue[i][1],eventTime] ).start()
	while schedTime.total_seconds() > 0:
		ntime = datetime.now()
		schedTime = (eventTime - ntime)
		print "Time until next event: " + str(schedTime)
		time.sleep(1)
	return(timedelta(hours=0))		
		
timer = main()
while 1:
	Timer(timer.total_seconds(),main,()).start

