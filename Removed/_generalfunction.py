import os

from openpyxl import load_workbook
from datetime import datetime
import time
import imutils
cwd = os.path.dirname(os.path.realpath(__file__))

################################################################################################################
# Sleep for an amount of time
# total_miliseconds
def Sleep(total_miliseconds):
	time.sleep(total_miliseconds/1000)

def basic_tap(Device, )

#Tap on Location_Object
def Tap(Device, Location_Object):
	command = "input tap " + str(Location_Object['x']) + " " + str(Location_Object['y'])
	Device.shell(command)
	return



#Swipe from A -> B
def Swipe(Device, Location_Object_A, Location_Object_B):
	command = "input swipe " + str(Location_Object_A['x']) + " " + str(Location_Object_A['y']) + " " + str(Location_Object_B['x']) + " " + str(Location_Object_B['y'])
	print(command)
	Device.shell(command)
	return

#Swipe up
def Swipe_Up(Device, Location_Object_A, Range):

	Location_Object_B = {'x': Location_Object_A['x'], 'y': Location_Object_A['y'] - Range} 
	print('Swipe from', Location_Object_A, 'to', Location_Object_B)
	command = "input swipe " + str(Location_Object_A['x']) + " " + str(Location_Object_A['y']) + " " + str(Location_Object_B['x']) + " " + str(Location_Object_B['y'])
	print(command)
	Device.shell(command)
	return

def Send_Text(Device, Text):
	command = "input text \' %s \'" %Text
	Device.shell(command)
	return

