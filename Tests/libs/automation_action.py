

from ppadb.client import Client as AdbClient
import os, sys
import cv2
import numpy as np
import time
#import imutils

#cwd = os.path.dirname(os.path.realpath(__file__))
CWD = os.path.abspath(os.path.dirname(sys.argv[0]))

Screenshot  = CWD + "\\"

if not os.path.isdir(Screenshot):
	try:
		os.mkdir(Screenshot)
	except OSError:
		print ("Creation of the directory %s failed" % Screenshot)

# Return result format:
# {
# 	"Type": Execute/Result/..,
#	"Status": True/False,
#	"Details": {},
#	"Screenshot": Array,
# }

class Automation_Action:
	# Serial: Device's serial
	# DB: Database's Path.
	def __init__(self):
		self.action_list = []
		

	def append_action_list(self, type = None, name = None, argument = [], description = ''):
		
		assert (type  == None), "Type is not decalred"
		assert (name  == None), "Name is not decalred"

		_action = {}
		_action['type'] = type
		_action['name'] = name
		_action['arg'] = argument
		_action['description'] = description

		self.action_list.append(_action)
		return _action

	