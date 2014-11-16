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
event_types = ["MACRO","PI"]

#inputs
cal = raw_input("Calender: (google or office) ")
deltime = timedelta(minutes=60)
#filename = raw_input("Enter macro file: ")
#event_name = "MACRO:" + filename

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
	def click(x,y):
		pythoncom.CoInitialize()
		win32api.SetCursorPos((x,y))
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,x,y,0,0)
		win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,x,y,0,0)
	def run_nt(event):
		pythoncom.CoInitialize()
		filestuff=open(filename,'r')
		times_click = []
		times_char = []
		xandy = []
		chars = []
		for line in filestuff:
			char=None
			temp = re.split('\s',line)
			if temp[0] == 'click':
				x_loc=temp[1]
				y_loc=temp[2]
				times_click.append(float(temp[3])) #time (ms)
				xandy.append([int(x_loc),int(y_loc)])
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
			s.enter(times_char[i],1,outkey,(chars[i]))
			#Timer(times_char[i],outkey,args=[chars[i]]).start()
		for i in range(len(times_click)):
			print times_click[i]
			print xandy[i]
			s.enter(times_click[i],1,click,(xandy[i][0], xandy[i][1]))
			#Timer(times_click[i],click,args=[xandy[i][0], xandy[i][1]]).start()
		s.run()
		    #queue.enter(delay=time_start,0,action=outkey,arg=(char))
	#End Alex's Code
else:
	#Pi code goes here!!!!
	def run(event,options,etime):
		print "Executing: ", event, "\nTime: ", etime
		if event == "webpage":
			webbrowser.open(options)

if(cal=="google"):
	from oauth2client.client import OAuth2WebServerFlow
	from oauth2client.file import Storage
	from oauth2client.tools import run_flow
	calID = raw_input("Enter google calendar ID: ")
	def main():
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
			event_name = re.split(':',re.split('\'',temp[i])[3])
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
		if schedTime > timedelta(hours=0):
			print("Time until next event: ", schedTime)
		else: 
			print("No events detected today, sleeping for 24hrs")
			return(timedelta(hours=24))
		run_command(schedTime,command_queue)
		return(timedelta(hours=0))
	timer = main()
	while 1:
		Timer(timer.total_seconds(),main,()).start
	
elif(cal=="office"):
	import requests
	import getpass
	def main():
		EMAIL = raw_input("Username: ")
		PASSWORD = getpass.getpass("Password: ")
		
		response = requests.get('https://outlook.office365.com/EWS/OData/users/'+EMAIL+'/Events', auth=(EMAIL, PASSWORD))
		print response.json()
		ntime = datetime.now()
		for event in response.json()['value']:
  			print(event['Subject'])
  			eventTime = datetime.strptime(event['Start'], "%Y-%m-%dT%H:%M:%SZ") - timedelta(hours=5) 
  			print eventTime
  			if (eventTime-ntime) < timedelta(hours=1) and (eventTime-ntime) > timedelta(hours=0):
  				break
  		schedTime = eventTime - ntime
		if schedTime > timedelta(hours=0):
			print("Time until next event: ", schedTime)
		else: 
			print("No events detected today, sleeping for 24hrs")
			return(timedelta(hours=24))
		run_command(schedTime,command_queue)
		return(timedelta(hours=0))
	timer = main()
	while 1:
		Timer(timer.total_seconds(),main,()).start
def run_command(schedTime,command_queue):
	if os.name == 'nt':
		Timer(schedTime.total_seconds(), run_nt,args=[command_queue[i][1]] ).start()
	else:
		Timer(schedTime.total_seconds(), run,args=[command_queue[i][1],eventTime,command_queue[i][2]] ).start()
	while schedTime.total_seconds() > 0:
		ntime = datetime.now()
		schedTime = (eventTime - ntime)
		print "Time until next event: " + str(schedTime)
		time.sleep(1)
