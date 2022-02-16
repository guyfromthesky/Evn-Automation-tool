#System variable and io handling
import sys
import os
import multiprocessing
from multiprocessing import Process , Queue, Manager
import queue 
import subprocess
#Get timestamp
import time
from datetime import datetime
import configparser
import random

#GUI
from tkinter import *
from tkinter.ttk import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from tkinter import colorchooser
from tkinter import scrolledtext 

#from tkinter import style

from openpyxl import load_workbook


from general_function import *
# Function use for the action builder here
from automation_driver import Automation as Tester
from ppadb.client import Client as AdbClient

from libs.configmanager import ConfigLoader
from libs.version import get_version
from libs.tkinter_extension import AutocompleteCombobox
import cv2
import pytesseract


CWD = os.path.abspath(os.path.dirname(sys.argv[0]))
ADBPATH = '\"' + CWD + '\\adb\\adb.exe' + '\"'
#MyTranslatorAgent = 'google'
Tool = "Automation Execuser"
VerNum = '0.3.0d'
version = Tool  + " " +  VerNum
DELAY1 = 20
DELAY2 = 100

#**********************************************************************************
# UI handle ***********************************************************************
#**********************************************************************************

class Automation_Execuser(Frame):
	def __init__(self, Root, Queue = None, Manager = None,):
		
		Frame.__init__(self, Root) 
		#super().__init__()
		self.parent = Root 

		# Queue
		self.Process_Queue = Queue['Process_Queue']
		self.Result_Queue = Queue['Result_Queue']
		self.Status_Queue = Queue['Status_Queue']
		self.Debug_Queue = Queue['Debug_Queue']

		self.Manager = Manager['Default_Manager']

		self.Options = {}

		# UI Variable
		self.Button_Width_Full = 20
		self.Button_Width_Half = 15
		
		self.PadX_Half = 5
		self.PadX_Full = 10
		self.PadY_Half = 5
		self.PadY_Full = 10
		self.StatusLength = 120
		self.AppLanguage = 'en'

		self.DB_Path = ""
		self.Test_Case_Path = ""
		self.TestCase = None


		self.App_LanguagePack = {}

		self.init_App_Setting()

		if self.AppLanguage != 'kr':
			from libs.languagepack import LanguagePackEN as LanguagePack
		else:
			from libs.languagepack import LanguagePackKR as LanguagePack

		self.LanguagePack = LanguagePack

		# Init function

		self.parent.resizable(False, False)
		self.parent.title(version)
		# Creating Menubar 
		
		#**************New row#**************#
		self.Notice = StringVar()
		self.Debug = StringVar()
		self.Progress = StringVar()
	
		#Generate UI
		self.Generate_Menu_UI()
		self.Generate_Tab_UI()
		self.init_UI()
		self.init_UI_Data()
		'''
		try:
			print('Start server')
			os.popen( ADBPATH + ' start-server')
			app_name = 'android_touch'
			print('Get CPU profile')
			#self.CPU = os.popen(ADBPATH + ' shell getprop ro.product.cpu.abi').read()
			process = subprocess.Popen(ADBPATH + ' shell getprop ro.product.cpu.abi', stdout=subprocess.PIPE, stderr=None, shell=True)
			CPU = process.communicate()[0].decode("utf-8") 
			CPUFamily = CPU.replace('\r\n', "")
			print('CPU family: ', CPUFamily)
			self.port = 0
			if CPUFamily != "":
				print('Push touch to device')
				if (os.path.isfile(CWD + '\\libs\\arm64-v8a\\touch')):
					print('file available')
				str_command = '%s push \"%s\\libs\\%s\\touch\" /data/local/tmp' % (ADBPATH, CWD, CPUFamily)
				print('Command:', str_command)
				os.system(str_command)
				print('Launch touch on device')
				os.system(ADBPATH + ' shell chmod 755 /data/local/tmp/touch') 
				os.system(ADBPATH + ' shell /data/local/tmp/touch')
				print('Looking for port')
				for port in range(50000,65535):
					self.port = port
					process = subprocess.Popen(ADBPATH + ' forward tcp:8080 tcp:' + str(self.port), stdout=subprocess.PIPE, stderr=None, shell=True)
					return_message = process.communicate()
					for message in return_message:
						if message != None:
							str_message = message.decode("utf-8") 
							if str_message == self.port:
								break
		 
					print('str_message', str_message)
					return_port = str_message.replace('\r\n', "")
					print('return_port', return_port)
					if return_port == self.port:
						break
			print('Current port:', self.port)

		except Exception as e:
			print('Error:', e)
		'''
		
		print('Get device serial')
		self.Get_Serial()
		self.after(DELAY2, self.status_listening)

	# UI init
	def init_UI(self):
	
		self.Generate_Automation_Execution_UI(self.Main)
		self.Generate_OCR_Setting_UI(self.OCR_SETTING)
		self.Generate_Debugger_UI(self.Controller)
		#self.Generate_Folder_Comparision_UI(self.FolderComparison)
		#self.Generate_Optimizer_UI(self.Optimizer)
		#self.Generate_Debugger_UI(self.Process)
		
		# Debugger

	#UI Function

	def Generate_Menu_UI(self):
		menubar = Menu(self.parent) 
		# Adding File Menu and commands 
		file = Menu(menubar, tearoff = 0)
		# Adding Help Menu
		help_ = Menu(menubar, tearoff = 0) 
		menubar.add_cascade(label =  self.LanguagePack.Menu['Help'], menu = help_) 
		help_.add_command(label =  self.LanguagePack.Menu['GuideLine'], command = self.Menu_Function_Open_Main_Guideline) 
		help_.add_separator()
		help_.add_command(label =  self.LanguagePack.Menu['About'], command = self.Menu_Function_About) 
		self.parent.config(menu = menubar)

		# Adding Help Menu
		language = Menu(menubar, tearoff = 0) 
		menubar.add_cascade(label =  self.LanguagePack.Menu['Language'], menu = language) 
		language.add_command(label =  self.LanguagePack.Menu['Hangul'], command = self.SetLanguageKorean) 
		language.add_command(label =  self.LanguagePack.Menu['English'], command = self.SetLanguageEnglish) 
		self.parent.config(menu = menubar) 	

	def Generate_Tab_UI(self):
		self.TAB_CONTROL = ttk.Notebook(self.parent)
		#Tab
		
		self.Main = ttk.Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.Main, text= 'Main')
		
		self.OCR_SETTING = Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.OCR_SETTING, text= 'OCR Setting')
		#Tab
		self.Controller = ttk.Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.Controller, text= 'Controller')
		#Tab
		

		'''
		self.FolderComparison = ttk.Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.FolderComparison, text= self.LanguagePack.Tab['FolderComparison'])
		#Tab
		self.Optimizer = ttk.Frame(self.TAB_CONTROL)
		self.TAB_CONTROL.add(self.Optimizer, text= self.LanguagePack.Tab['Optimizer'])
		'''
		#Tab
		#self.Process = ttk.Frame(self.TAB_CONTROL)
		#self.TAB_CONTROL.add(self.Process, text= self.LanguagePack.Tab['Debug'])
		

		self.TAB_CONTROL.pack(expand=1, fill="both")
		return

	def Generate_Automation_Execution_UI(self, Tab):
		
		Row = 1

		#self.Str_Test_Case_Path.set( CWD + '\\Testcase\\Sample_Automation_Testcase.xlsx')
		Label(Tab, text=  self.LanguagePack.Label['TestCaseList']).grid(row=Row, column=1, columnspan=2, padx=5, pady=5, sticky= W)
		self.Entry_New_File_Path = Entry(Tab,width = 80, state="readonly", textvariable=self.Str_Test_Case_Path)
		self.Entry_New_File_Path.grid(row=Row, column=3, columnspan=7, padx=4, pady=5, sticky=W+E)
		Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Browse_Test_Case_File).grid(row=Row, column=10, padx=0, pady=0, sticky=W)
		#Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['SelectFontColor'], command= self.Btn_Select_Font_Colour).grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)
		
	
		Row += 1
		Label(Tab, text= 'Execute List').grid(row=Row, column=1, columnspan=2, padx=5, pady=5, sticky=W)
		self.ExecuteList = AutocompleteCombobox(Tab)

		#self.TestProject = AutocompleteCombobox(Tab)
		self.ExecuteList.set_completion_list([])
		self.ExecuteList.Set_Entry_Width(20)
		self.ExecuteList.grid(row=Row, column=3, padx=5, pady=5, sticky=W)

		Label(Tab, text=self.LanguagePack.Label['Serial']).grid(row=Row, column=7, padx=5, pady=5, sticky=W)			
		self.TextSerial = AutocompleteCombobox(Tab)
		self.TextSerial.Set_Entry_Width(20)
		
		self.TextSerial.grid(row=Row, column=8, columnspan=2,padx=5, pady=5, sticky=W+E)
		self.TextSerial.set_completion_list([])
		Button(Tab, width = self.Button_Width_Half, text=  "Get Device", command= self.Get_Serial).grid(row=Row, column=10,padx=0, pady=0, sticky=W)
		
		'''
		Row += 1
		Label(Tab, text=self.LanguagePack.Label['TestFeature']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.TestFeature = AutocompleteCombobox(Tab)
		self.TestFeature.set_completion_list(['Gacha'])
		self.TestFeature.Set_Entry_Width(30)
		self.TestFeature.grid(row=Row, column=3, padx=5, pady=5, sticky=W)
		'''
		
		#Button(Tab, width = self.Button_Width_Half, text=  'Get Test info', command= self.Btn_Generate_TestCase).grid(row=Row, column=10,padx=0, pady=0, sticky=W)

		Row += 1
		Label(Tab, text= self.LanguagePack.Label['WorkingLang']).grid(row=Row, column=1, columnspan=2, padx=5, pady=5, sticky=W)
		self.option_working_language = AutocompleteCombobox(Tab)
		self.option_working_language.Set_Entry_Width(20)
		self.option_working_language.grid(row=Row, column=3, padx=5, pady=5, sticky=W)

		Label(Tab, text= "Device IP").grid(row=Row, column=7, padx=5, pady=5, sticky=W)	
		self.Device_IP = Text(Tab, width=30, height=1, undo=True, wrap=WORD)
		self.Device_IP.grid(row=Row, column=8, columnspan=2, padx=5, pady=5, sticky=E)
		self.Device_IP.insert("end", "0.0.0.0")
		Button(Tab, width = self.Button_Width_Half, text=  "Wireless Connect", command= self.Connect_Device).grid(row=Row, column=10, padx=0, pady=0, sticky=W)

		Row+=1
		Treeview_Row = 20
		self.Generate_Treeview_Advanced_UI(Tab, Row, Treeview_Row)

		Row+=1
		Button(Tab, width = self.Button_Width_Half, text=  'Add above', command= self.insert_treeview_above).grid(row=Row, column=10,padx=0, pady=0, sticky=W)

		Row+=1
		Button(Tab, width = self.Button_Width_Half, text=  "Add", command= self.add_treeview_row).grid(row=Row, column=10,padx=0, pady=0, sticky=W)

		Row+=1
		Button(Tab, width = self.Button_Width_Half, text=  "Add below", command= self.insert_treeview_below).grid(row=Row, column=10,padx=0, pady=0, sticky=W)

		Row+=1
		Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Stop'], command= self.Stop).grid(row=Row, column=10,padx=0, pady=0, sticky=W)	
		Row+=1
		self.Btn_Execute = Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Execute'], command= self.Btn_Execute_Script)
		self.Btn_Execute.grid(row=Row, column=10,padx=0, pady=0, sticky=W)


		Row += Treeview_Row
		self.Debugger = scrolledtext.ScrolledText(Tab, width=100, height=10, undo=True, wrap=WORD, )
		self.Debugger.grid(row=Row, column=1, columnspan=9, rowspan=10, padx=5, pady=5, sticky=W+E+N+S)
		Button(Tab, width = self.Button_Width_Half, text=  "Clear Log", command= self.ClearLog).grid(row=Row, column=10,padx=0, pady=0, sticky=W)

		#ScrollBar = Scrollbar(Tab, bg="green")
		#ScrollBar.pack( side = RIGHT, fill = Y )

	def Generate_Treeview_Advanced_UI(self, Tab, start_row, row_span):
		TreeView_Row = start_row
		TreeView_Size = row_span
		Treeview_Col = 9
		self.Treeview = Treeview(Tab)
		self.Focused_Item = None
		self.Treeview.grid(row=TreeView_Row, column=1, columnspan=Treeview_Col, rowspan=20, padx=5, pady=5, sticky = W+E)
		verscrlbar = Scrollbar(Tab, orient ="vertical", command = self.Treeview.yview)
		self.Treeview.configure( yscrollcommand=verscrlbar.set)
	
		self.Treeview.Scrollable = True
		self.Treeview['columns'] = ('Type', 'Action', 'Arg1', 'Arg2', 'Arg3', 'Arg4', 'Arg5', 'Arg6')

		self.Treeview.column('#0', width=0, stretch=NO)
		self.Treeview.heading('#0', text='', anchor=CENTER)

		for column in self.Treeview['columns']:
			self.Treeview.column(column, anchor=CENTER, width=100)
			self.Treeview.heading(column, text=column, anchor=CENTER)

		verscrlbar.grid(row=TreeView_Row, column=Treeview_Col, rowspan=TreeView_Size, padx=5, pady=5,  sticky = N+S+E)
		Tab.grid_columnconfigure(TreeView_Size, weight=0, pad=5)
		styles = Style()
		styles.configure('Treeview',rowheight=row_span)

		self.Treeview.bind("<Delete>", self.delete_treeview_line)	
		self.Treeview.bind("<Double-1>", self.Treeview_OCR_Select_Row)

	def Generate_OCR_Setting_UI(self, Tab):
		''''
		Create Setting Tab
		'''
		Row = 1
		Label(Tab, text= self.LanguagePack.Label['TesseractPath']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Text_TesseractPath = Entry(Tab,width = 100, state="readonly", textvariable=self.TesseractPath)
		self.Text_TesseractPath.grid(row=Row, column=3, columnspan=5, padx=5, pady=5, sticky=E+W)
		Button(Tab, width = self.Button_Width_Full, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Select_Tesseract_Path).grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=E)
		
		Row += 1
		Label(Tab, text= self.LanguagePack.Label['TesseractDataPath']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Text_TesseractDataPath = Entry(Tab,width = 100, state="readonly", textvariable=self.TesseractDataPath)
		self.Text_TesseractDataPath.grid(row=Row, column=3, columnspan=5, padx=5, pady=5, sticky=E+W)
		Button(Tab, width = self.Button_Width_Full, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Select_Tesseract_Data_Path).grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=E)
		
		Row += 1
		Label(Tab, text= self.LanguagePack.Label['DBPath']).grid(row=Row, column=1, padx=5, pady=5, sticky=W)
		self.Text_DB_Path = Entry(Tab,width = 100, state="readonly", textvariable=self.DBPath)
		self.Text_DB_Path.grid(row=Row, column=3, columnspan=5, padx=5, pady=5, sticky=E+W)
		Button(Tab, width = self.Button_Width_Full, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Browse_DB_File).grid(row=Row, column=9, columnspan=2, padx=5, pady=5, sticky=E)
		
	def Generate_Debugger_UI(self,Tab):
		Row = 1
		Button(Tab, width = self.Button_Width_Half, text=  "TAB", command= self.Btn_Send_Tab).grid(row=Row, column=1,padx=5, pady=5, sticky=W)
		Button(Tab, width = self.Button_Width_Half, text=  "Enter", command= self.Btn_Send_Enter).grid(row=Row, column=2,padx=5, pady=5, sticky=W)
		Button(Tab, width = self.Button_Width_Half, text=  "4 touch", command= self.Btn_Send_4_Touch).grid(row=Row, column=3,padx=5, pady=5, sticky=W)
		Row += 1
		Button(Tab, width = self.Button_Width_Half, text=  "Home", command= self.Btn_Send_Home).grid(row=Row, column=1,padx=5, pady=5, sticky=W)
		Button(Tab, width = self.Button_Width_Half, text=  "Backkey", command= self.Btn_Send_Backkey).grid(row=Row, column=2,padx=5, pady=5, sticky=W)

		Row += 1
		self.Str_Template_Path = StringVar()
		Label(Tab, text= 'Template path').grid(row=Row, column=1, padx=5, pady=5, sticky= W)
		self.Entry_Old_File_Path = Entry(Tab,width = 80, state="readonly", textvariable=self.Str_Template_Path)
		self.Entry_Old_File_Path.grid(row=Row, column=2, columnspan=5, padx=5, pady=5, sticky=W)
		Button(Tab, width = self.Button_Width_Half, text=  self.LanguagePack.Button['Browse'], command= self.Btn_Browse_Template_File).grid(row=Row, column=8, padx=5, pady=5, sticky=W)
		Button(Tab, width = self.Button_Width_Half, text= 'Tap', command= self.Btn_Tap_Template).grid(row=Row, column=9, columnspan=2,padx=5, pady=5, sticky=W)
		
###########################################################################################
# OCR Setting
###########################################################################################

	def Btn_Select_Tesseract_Path(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectDB'],filetypes = (("Executable files","*.exe" ), ), )	
		if os.path.isfile(filename):
			_tess_path = self.CorrectPath(filename)
			self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'AUTO_TOOL', 'tess_path', _tess_path, True)
			pytesseract.pytesseract.tesseract_cmd = _tess_path
			self.TesseractPath.set(_tess_path)
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['TessNotSelect'])

	def Btn_Select_Tesseract_Data_Path(self):
		folder_name = filedialog.askdirectory(title =  self.LanguagePack.ToolTips['SelectSource'],)	
		if os.path.isdir(folder_name):
			folder_name = self.CorrectPath(folder_name)
			self.TesseractDataPath.set(folder_name)

			self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'AUTO_TOOL', 'tess_data', folder_name, True)

			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + folder_name)
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])

	def Btn_Browse_DB_File(self):
			
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("Workbook files", "*.xlsx *.xlsm"), ), multiple = False)	
		if filename != "":
			self.DB_Path = self.CorrectPath(filename)
			self.Str_DB_Path.set(self.DB_Path)
			
			self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'AUTO_TOOL', 'db_path', self.DB_Path, True)
			self.Write_Debug(self.LanguagePack.ToolTips['DataSelected'] + ": " + self.DB_Path)
	
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return


	def OCR_Setting_Set_Scan_Type(self, scan_type):		
		self.ScanType.set(scan_type)
		self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'AUTO_TOOL', 'scan_type', scan_type)
		self.Write_Debug(self.LanguagePack.ToolTips['ScanTypeUpdate'] + str(scan_type) + '.')
		if scan_type == 'DB Create':
			self.Write_Debug(self.LanguagePack.ToolTips['DBCreate'])
		elif scan_type == 'Text only':
			self.Write_Debug(self.LanguagePack.ToolTips['TextScan'])
		elif scan_type == 'Image only':
			self.Write_Debug(self.LanguagePack.ToolTips['ImageScan'])
		elif scan_type == 'Image and Text':
			self.Write_Debug(self.LanguagePack.ToolTips['AdvancedScan'])
	



	def OCR_Setting_Set_Browse_Type(self):
		_browse_type = self.Browse_Type.get()
		if _browse_type == 1:
			_status = 'folder'
		else:
			_status = 'file'
		
		self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'OCR_TOOL', 'browsetype', _browse_type)

		self.Write_Debug(self.LanguagePack.ToolTips['BrowseTypeUpdate'] + str(_status))

	def OCR_Setting_Set_Working_Resolution(self):
		_resolution_index = self.Resolution.get()
		if _resolution_index == 1:
			self.WorkingResolution = 720
		else:
			self.WorkingResolution = 1080
		
		self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'OCR_TOOL', 'resolution', _resolution_index)

		self.Write_Debug(self.LanguagePack.ToolTips['SetResolution'] + str(self.WorkingResolution) + 'p')

	def OCR_Setting_Set_Working_Language(self, select_value):
		
		self.AppConfig.Save_Config(self.AppConfig.Auto_Tool_Config_Path, 'OCR_TOOL', 'scan_lang', select_value)
		
		self.Write_Debug(self.LanguagePack.ToolTips['SetScanLanguage'] + str(select_value))
	
	

###########################################################################################
# General functions
###########################################################################################

	def CorrectPath(self, path):
		if sys.platform.startswith('win'):
			return str(path).replace('/', '\\')
		else:
			return str(path).replace('\\', '//')
	
	def CorrectExt(self, path, ext):
		if path != None and ext != None:
			Outputdir = os.path.dirname(path)
			baseName = os.path.basename(path)
			sourcename = os.path.splitext(baseName)[0]
			newPath = self.CorrectPath(Outputdir + '/'+ sourcename + '.' + ext)
			return newPath

	def Write_Debug(self, text):
		'''
		Function write the text to debugger box and move to the end of the box
		'''
		self.Debugger.insert("end", "\n")
		self.Debugger.insert("end", str(text))

		self.Debugger.yview(END)		

	def entry_next(self, event):
		event.widget.tk_focusNext().focus()
		return("break")

###########################################################################################
# Treeview FUNCTION
###########################################################################################

	def delete_treeview_line(self, event):
		'''
		Function activate when select an entry from a Treeview and press Delete btn
		'''
		selected = self.Treeview.selection()
		to_remove = []
		for child_obj in selected:
			child = self.Treeview.item(child_obj)
			tm_index = child['values'][0]
			to_remove.append(tm_index)
			self.Treeview.delete(child_obj)

	# Obsoleted.
	def double_right_click_treeview(self, event):
		'''
		Function activate when double click an entry from Treeview
		'''
		focused = self.Treeview.focus()
		child = self.Treeview.item(focused)
		
		self.Debugger.insert("end", "\n")
		self.Debugger.insert("end", 'Korean: ' + str(child["text"]))
		self.Debugger.insert("end", "\n")
		self.Debugger.insert("end", 'English: ' + str(child["values"][0]))
		self.Debugger.yview(END)


	# Nam will check
	def load_tm_list(self):
		"""
		When clicking the [Load] button in TM Manager tab
		Display the pair languages in the text box.
		"""
		self.remove_treeview()
		
		_area_list = []

		for location in _area_list:	
			try:
				self.Treeview.insert('', 'end', text= '', values=( str(location['index']), str(location['x']), str(location['y']), str(location['h']), str(location['w'])))
			except:
				pass

	def add_treeview_row(self):
		'''
		Add a row to the current Treeview
		'''
		self.Treeview.insert('', 'end', text= '', values=(1,1,1,1,1,1,1,1))

	def insert_treeview_above(self):
		'''
		Add a row above the current Treeview
		'''
		index = self.Treeview.index(self.Treeview.focus())
		self.Treeview.insert('', index, text= '', values=(2,2,2,2,2,2,2,2,2))

	def insert_treeview_below(self):
		'''
		Add a row below the current Treeview
		'''
		index = self.Treeview.index(self.Treeview.focus())
		self.Treeview.insert('', index+1, text= '', values=(3,3,3,3,3,3,3,3))
		

	def remove_treeview(self):
		for i in self.Treeview.get_children():
			self.Treeview.delete(i)

	def Btn_OCR_Select_Area(self):
		
		if self.OCR_File_Path != None:
			_index = random.randint(0, len(self.OCR_File_Path)-1)
			if os.path.isfile(self.OCR_File_Path[_index]):
				im = cv2.imread(self.OCR_File_Path[_index])
				(_h, _w) = im.shape[:2]
				ratio = 720 / _h
				if ratio != 1:
					width = int(im.shape[1] * ratio)
					height = int(im.shape[0] * ratio)
					dim = (width, height)
					im = cv2.resize(im, dim, interpolation = cv2.INTER_AREA)

				for row in self.Treeview.get_children():
					child = self.Treeview.item(row)
					im = cv2.rectangle(im, (child["values"][0], child["values"][1]), (child["values"][0] + child["values"][2], child["values"][1] + child["values"][3]), (255,0,0), 2)

				location = cv2.selectROI("Sekect scan area", im, showCrosshair=False,fromCenter=False)
				cv2.destroyAllWindows() 
				self.Treeview.insert('', 'end', text= '', values=(str(location[0]), str(location[1]), str(location[2]), str(location[3])))
			else:
				self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])		
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])	
		
	def Btn_OCR_Select_Area_Advanced(self):
		_scan_type = self.ScanType.get()

		if self.OCR_File_Path != None:
			_index = random.randint(0, len(self.OCR_File_Path)-1)
			if os.path.isfile(self.OCR_File_Path[_index]):
				im = cv2.imread(self.OCR_File_Path[_index])
				(_h, _w) = im.shape[:2]
				ratio = 720 / _h
				if ratio != 1:
					width = int(im.shape[1] * ratio)
					height = int(im.shape[0] * ratio)
					dim = (width, height)
					im = cv2.resize(im, dim, interpolation = cv2.INTER_AREA)

				for row in self.Treeview.get_children():
					child = self.Treeview.item(row)
					if _scan_type in ['Image and Text', 'DB Create', 'Text only']:
						im = cv2.rectangle(im, (child["values"][0], child["values"][1]), (child["values"][0] + child["values"][2], child["values"][1] + child["values"][3]), (255,0,0), 2)
					if _scan_type in ['Image and Text', 'DB Create', 'Image only']:
						im = cv2.rectangle(im, (child["values"][4], child["values"][5]), (child["values"][4] + child["values"][6], child["values"][5] + child["values"][7]), (255,255,0), 2)
				location = [0,0,0,0]
				location2 = [0,0,0,0]
				if _scan_type in ['Image and Text', 'DB Create', 'Text only']:
					location = cv2.selectROI("Select TEXT area", im, showCrosshair=False,fromCenter=False)
					im = cv2.rectangle(im, (location[0], location[1]), (location[0] + location[2], location[1] + location[3]), (255,0,0), 2)
					cv2.destroyAllWindows() 
				if _scan_type in ['Image and Text', 'DB Create', 'Image only']:
					location2 = cv2.selectROI("Select COMPONENT area", im, showCrosshair=False,fromCenter=False)
					cv2.destroyAllWindows() 
				
				self.Treeview.insert('', 'end', text= '', values=(str(location[0]), str(location[1]), str(location[2]), str(location[3]), str(location2[0]), str(location2[1]), str(location2[2]), str(location2[3]) ))
				
			else:
				self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])		
		else:
			self.Write_Debug(self.LanguagePack.ToolTips['SourceDocumentEmpty'])				

	def Btn_OCR_Input_Text_Area(self):

		self.Focused_Item = self.Treeview.focus()
		child = self.Treeview.item(self.Focused_Item)
		
		

		_x = self.Str_CenterX.get("1.0", END).replace('\n', '')
		if _x == '': _x = 0
		_y = self.Str_CenterY.get("1.0", END).replace('\n', '')
		if _y == '': _y = 0
		_w = self.Str_Weight.get("1.0", END).replace('\n', '')
		if _w == '': _w = 0
		_h = self.Str_Height.get("1.0", END).replace('\n', '')
		if _h == '': _h = 0
		
		if self.Region_Type.get() == 1:
			self.Treeview.insert('', 'end', text= '', values=(str(int(_x)), str(int(_y)), str(int(_w)), str(int(_h)), child["values"][4], child["values"][5], child["values"][6], child["values"][7]))
		else:
			self.Treeview.insert('', 'end', text= '', values=(child["values"][0], child["values"][1], child["values"][2], child["values"][3], str(int(_x)), str(int(_y)), str(int(_w)), str(int(_h))))

		#self.Update_Treeview_Advanced_UI()

	def Treeview_OCR_Select_Row(self, event):
		'''
		Function activate when double click an entry from Treeview
		'''
		self.Focused_Item = self.Treeview.focus()
		child = self.Treeview.item(self.Focused_Item)
		index = self.Treeview.index(self.Focused_Item)
		print('index number:', index)
		

	def Btn_OCR_Update_Area(self):

		if self.Focused_Item != None:
			child = self.Treeview.item(self.Focused_Item)
			_x = self.Str_CenterX.get("1.0", END).replace('\n', '')
			if _x == '': _x = 0
			_y = self.Str_CenterY.get("1.0", END).replace('\n', '')
			if _y == '': _y = 0
			_w = self.Str_Weight.get("1.0", END).replace('\n', '')
			if _w == '': _w = 0
			_h = self.Str_Height.get("1.0", END).replace('\n', '')
			if _h == '': _h = 0
			
			if self.Region_Type.get() == 1:
				self.Treeview.item(self.Focused_Item, text="", values=(str(int(_x)), str(int(_y)), str(int(_w)), str(int(_h)), child["values"][4], child["values"][5], child["values"][6], child["values"][7]))
			else:
				self.Treeview.item(self.Focused_Item, text="", values=(child["values"][0], child["values"][1], child["values"][2], child["values"][3], str(int(_x)), str(int(_y)), str(int(_w)), str(int(_h))))


			#self.Treeview.item(self.Focused_Item, text="", values=(child["values"]))
			self.Focused_Item = None
			

	# Background function
	def Get_Serial(self):
		try:
			client = AdbClient(host="127.0.0.1", port=5037)
			devices = client.devices()
			Serrial = []
			for device in devices:
				Serrial.append(device.serial)
			
			self.TextSerial.set_completion_list(Serrial)
			self.TextSerial.current(0)
		except:	
			pass

	def Connect_Device(self):
		IP = self.Device_IP.get("1.0", END).replace('\n', '')
		os.popen( ADBPATH + ' tcpip 5555')
		os.popen( ADBPATH + ' connect ' + str(IP) + ':5555')
		self.Get_Serial()	

	def Btn_Send_Tab(self):
		os.popen( ADBPATH + ' shell input keyevent \'61\'')
	
	def Btn_Send_Enter(self):
		os.popen( ADBPATH + ' shell input keyevent \'66\'')	

	def Btn_Send_4_Touch(self):
		os.system( ADBPATH + ' forward tcp:9889 tcp:9889')
		Four_Touch()

	def Btn_Send_Backkey(self):
		os.popen( ADBPATH + ' shell input keyevent \'4\'')	

	def Btn_Send_Home(self):
		os.popen( ADBPATH + ' shell input keyevent \'3\'')		

	def Btn_Tap_Template(self):
		if self.Template_Path != None:
			self.Get_Serial()
			Serial = self.TextSerial.get().replace('\n','')
			AutoTester = Tester(self.Status_Queue, Serial, self.DB_Path)
			AutoTester.Tap_Template(self.Template_Path)

	def Update_Execute_List(self):
		print('Update execution list.')
		try:
			print(self.TestCase['Data'])
			self.ExecuteList.set_completion_list(self.TestCase['Data'])
			self.ExecuteList.current(0)
		except:
			pass	

	# Other function
	def Stop(self):
		try:
			if self.Automation_Processor.is_alive():
				self.Automation_Processor.terminate()
		except:
			pass
		self.Show_Error_Message('Test is terminated')


	# Other function
	def Btn_OCR_Update_Working_Language(self):
		_data_ = str(self.TesseractDataPath.get())
		_exe_ = str(self.TesseractPath.get())
		_tessdata_dir_config = '--tessdata-dir ' + "\"" + _data_ + "\""
		pytesseract.pytesseract.tesseract_cmd = _exe_
		#self.language_list = pytesseract.get_languages(config=_tessdata_dir_config)
		try:
			self.language_list = pytesseract.get_languages(config=_tessdata_dir_config)
			self.Write_Debug('Supported language list has been updated!')

		except Exception as e:
			self.Write_Debug('Tess path: ' + str(_exe_))
			self.Write_Debug('Data path: ' + str(_data_))
			self.Write_Debug('Error while updating supported language: ' + str(e))
			self.language_list = ['']

		self.option_working_language.set_completion_list(self.language_list)

	def ClearLog(self):
		self.Debugger.delete('1.0', END)
		return

	def Get_Test_Info(self):
		Test_Case = self.Str_Test_Case_Path.get()
		All = Function_Import_TestCase(Test_Case)
		TestInfo = All['Info']
		for Key in TestInfo:
			self.Debugger.insert("end", "\n\r")
			self.Debugger.insert("end", str(Key) + ':' + str(TestInfo[Key]))
			self.Debugger.yview(END)	



	def ImportTestCase(self, Test_Case_File_Path):
		print('Loading My Dictionary')
		if Test_Case_File_Path != None:
			if (os.path.isfile(Test_Case_File_Path)):
				xlsx = load_workbook(Test_Case_File_Path)
				DictList = []
				Dict = []
				for sheet in xlsx:
					sheetname = sheet.title.lower()
					if sheetname not in self.SpecialSheets:	
						EN_Coll = ""
						KR_Coll = ""
						database = None
						ws = xlsx[sheet.title]
						for row in ws.iter_rows():
							for cell in row:
								if cell.value == "KO":
									KR_Coll = cell.column_letter
									KR_Row = cell.row
									database = ws
								elif cell.value == "EN":
									EN_Coll = cell.column_letter
								if KR_Coll != "" and EN_Coll != "":
									DictList .append(sheet.title)
									break	
							if database!=  None:
								break		
						if database != None:
							for i in range(KR_Row, database.max_row): 
								KRAddress = KR_Coll + str(i+1)
								ENAddress = EN_Coll + str(i+1)
								#print('KRAddress', KRAddress)
								#print('ENAddress', ENAddress)
								KRCell = database[KRAddress]
								KRValue = KRCell.value
								ENCell = database[ENAddress]
								ENValue = ENCell.value
								if KRValue == None or ENValue == None or KRValue == 'KO' or ENValue == 'EN':
									continue
								elif KRValue != None and ENValue != None:
									Dict.append([KRValue, ENValue.lower()])
				print("Successfully load dictionary from: ", DictList)
				return Dict
			else:
				return([])	
		else:
			return([])	

	# Menu Function
	def Menu_Function_About(self):
		messagebox.showinfo("About....", "Creator: Evan")

	def Show_Error_Message(self, ErrorText):
		messagebox.showinfo('Error...', ErrorText)	

	def SaveAppLanguage(self, language):

		self.Notice.set('Update app language...') 

		config = configparser.ConfigParser()
		config.read(self.AppConfig)
		if not config.has_section('DocumentToolkit'):
			config.add_section('DocumentToolkit')
			cfg = config['DocumentToolkit']	
		else:
			cfg = config['DocumentToolkit']

		cfg['applang']= language
		with open(self.AppConfig, 'w') as configfile:
			config.write(configfile)
		self.Notice.set('Config saved...')
		return

	def SetLanguageKorean(self):
		self.AppLanguage = 'kr'
		self.SaveAppLanguage(self.AppLanguage)
		#self.initUI()
	
	def SetLanguageEnglish(self):
		self.AppLanguage = 'en'
		self.SaveAppLanguage(self.AppLanguage)
		#self.initUI()

	def Menu_Function_Open_Main_Guideline(self):
		#webbrowser.open_new(r"https://confluence.nexon.com/pages/viewpage.action?pageId=298119695")
		print('Done')

	def Function_Correct_Path(self, path):
		#print("path", path)
		return str(path).replace('/', '\\')
	
	def Function_Correct_EXT(self, path, ext):
		if path != None and ext != None:
			Outputdir = os.path.dirname(path)
			baseName = os.path.basename(path)
			sourcename, Obs_ext = os.path.splitext(baseName)
			newPath = self.CorrectPath(Outputdir + '/'+ sourcename + '.' + ext)
			return newPath

	def ErrorMsg(self, ErrorText):
		messagebox.showinfo('Error...', ErrorText)	

	def OpenOutput(self):
		Source = self.ListFile[0]

		Outputdir = os.path.dirname(Source)
		BasePath = str(os.path.abspath(Outputdir))
		subprocess.Popen('explorer ' + BasePath)
		
	def BtnLoadDocument(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("All type files","*.docx *.xlsx *.pptx *.msg"), ("Workbook files","*.xlsx *.xlsm"), ("Document files","*.docx"), ("Presentation files","*.pptx"), ("Email files","*.msg")), multiple = True)	
		if filename != "":
			self.ListFile = list(filename)
			self.CurrentSourceFile.set(str(self.ListFile[0]))
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def BtnLoadRawSource(self):
		#filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("Workbook files","*.xlsx *.xlsm *xlsb"), ("Document files","*.docx")), multiple = True)	
		FolderName = filedialog.askdirectory(title =  self.LanguagePack.ToolTips['SelectSource'])	
		if FolderName != "":
			self.RawFile = FolderName
			self.RawSource.set(str(FolderName))
			
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return
		'''
		if filename != "":
			self.RawFile = list(filename)
			self.RawSource.set(self.RawFile)
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return
		'''

	def BtnLoadTrackingSource(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("TM file","*.pkl"),),)	
		if filename != "":
			self.TrackingFile = self.CorrectPath(filename)
			print(self.TrackingFile)
			self.TrackingSource.set(self.CorrectPath(self.TrackingFile))
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def BtnLoadRawTM(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("TM file","*.pkl"),), multiple = True)	
		if filename != "":
			self.RawTMFile = list(filename)
			Display = self.CorrectPath(self.RawTMFile[0])
			self.RawTMSource.set(Display)
			
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return
	
	
	def BtnSelectColour(self):
		colorStr, self.BackgroundColor = colorchooser.askcolor(parent=self, title='Select Colour')
		
		
		if self.BackgroundColor == None:
			self.ErrorMsg('Set colour as defalt colour (Yellow)')
			self.BackgroundColor = 'ffff00'
		else:
			self.BackgroundColor = self.BackgroundColor.replace('#', '')
		#print(colorStr)
		#print(self.BackgroundColor)
		return

	def Btn_Browse_Old_Data_Folder(self):
		FolderName = filedialog.askdirectory(title =  self.LanguagePack.ToolTips['SelectSource'])	
		if FolderName != "":
			self.OldDataTable = FolderName
			self.OldDataString.set(str(FolderName))
			
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def Btn_Browse_New_Data_Folder(self):
		FolderName = filedialog.askdirectory(title =  self.LanguagePack.ToolTips['SelectSource'])	
		if FolderName != "":
			self.NewDataTable = FolderName
			
			self.NewDataString.set(str(FolderName))
			
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return	

	def BtnLoadDataLookupSource(self):
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("Workbook files *.xlsx"), ), multiple = True)	
		if filename != "":
			self.RawFile = list(filename)
			self.RawSource.set(self.RawFile)
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def onExit(self):
		self.quit()


	def init_App_Setting(self):
		
		self.Str_DB_Path = StringVar()
		self.Str_Test_Case_Path = StringVar()

		self.TesseractPath = StringVar()
		self.TesseractDataPath = StringVar()
		self.WorkingLanguage = StringVar()
		self.language_list = ['']

		self.DBPath = StringVar()

		self.Browse_Type = IntVar()

		self.Resolution = IntVar()
		self.CurrentDataSource = StringVar()

		self.ScanType = StringVar()

		self.Notice = StringVar()

		self.AppConfig = ConfigLoader('AUTO_TOOL')

		self.Configuration = self.AppConfig.Config
		self.AppLanguage  = self.Configuration['AUTO_TOOL']['app_lang']

		_tesseract_path = self.Configuration['AUTO_TOOL']['tess_path']
		pytesseract.pytesseract.tesseract_cmd = str(_tesseract_path)
		self.TesseractPath.set(_tesseract_path)

		_tesseract_data_path = self.Configuration['AUTO_TOOL']['tess_data']
		self.TesseractDataPath.set(_tesseract_data_path)

		_db_path = self.Configuration['AUTO_TOOL']['db_path']
		self.DBPath.set(_db_path)

		_browse_type = self.Configuration['AUTO_TOOL']['browsetype']
		self.Browse_Type.set(_browse_type)

		_resolution = self.Configuration['AUTO_TOOL']['resolution']
		self.Resolution.set(_resolution)

	def init_UI_Data(self):
		
		self.Btn_OCR_Update_Working_Language()
		_working_language = self.Configuration['AUTO_TOOL']['scan_lang']
		self.option_working_language.set(_working_language)

		_scan_type = self.Configuration['AUTO_TOOL']['scan_type']
		self.ScanType.set(_scan_type)
	
	def Btn_Select_Font_Colour(self):
		colorStr, self.Font_Color = colorchooser.askcolor(parent=self, title='Select Colour')
		
		
		if self.Font_Color == None:
			self.Error('Set colour as defalt colour (Yellow)')
			self.Font_Color = 'FF0000'
		else:
			self.Font_Color = self.Font_Color.replace('#', '')
		#print(colorStr)
		#print(self.BackgroundColor)
		return

	
	def Btn_Browse_Template_File(self):
			
		filename = filedialog.askopenfilename(title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("Template files", "*.jpg *.png"), ), multiple = False)	
		if filename != "":
			self.Template_Path = self.Function_Correct_Path(filename)
			self.Str_Template_Path.set(filename)
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def Btn_Browse_Test_Case_File(self):
		
		filename = filedialog.askopenfilename(initialdir = CWD + "//Testcase", title =  self.LanguagePack.ToolTips['SelectSource'],filetypes = (("Workbook files", "*.xlsx *.xlsm"), ), multiple = False)
		#cwd
		if filename != "":
			self.Test_Case_Path = self.Function_Correct_Path(filename)
			self.Str_Test_Case_Path.set(self.Test_Case_Path)
			self.Notice.set(self.LanguagePack.ToolTips['SourceSelected'])
		else:
			self.Notice.set(self.LanguagePack.ToolTips['SourceDocumentEmpty'])
		return

	def Btn_Execute_Script(self):
	
		self.Btn_Generate_TestCase()

		DB = self.Str_DB_Path.get()
		Execute_Value = self.ExecuteList.get().replace('\n','')
		Test_Case = self.Str_Test_Case_Path.get()
		Serial = self.TextSerial.get().replace('\n','')
		#MyDB = self.Function_Import_DB(self.DB_Path)
		try:
			self.Automation_Processor.terminate()
		except Exception as e:
			pass
		self.Automation_Processor = Process(target=Function_Execute_Script, args=(self.Status_Queue, self.Result_Queue, Serial, DB, Test_Case, self.TestCase, Execute_Value,))
		#Status_Queue, Result_Queue, Serial_Nummber, DB_Path, Test_Case_Path, TestCaseObject = []
		#self.Data_Compare_Process = Process(target=Old_Function_Compare_Excel, args=(self.Status_Queue, self.Process_Queue, Old_File, New_File, Output_Result, Sheet_Name, Index_Col, self.Background_Color, self.Font_Color,))
		self.Automation_Processor.start()
		self.after(DELAY1, self.Wait_For_Automation_Processor)	

	def status_listening(self):
		while True:
			try:
				Status = self.Status_Queue.get(0)
				if Status != None:
					self.Debugger.insert("end", "\n\r")
					ct = datetime.now()
					self.Debugger.insert("end", str(ct) + ": " + Status)
					self.Debugger.yview(END)
			except queue.Empty:
				break
		#self.check_device_connection()
		self.after(DELAY2, self.status_listening)

	def Wait_For_Automation_Processor(self):
		if (self.Automation_Processor.is_alive()):	
			self.after(DELAY1, self.Wait_For_Automation_Processor)
		else:
			self.Automation_Processor.terminate()
			#self.Show_Error_Message('Test is completed')

	def Btn_Generate_TestCase(self):
		DB = self.Str_DB_Path.get()
		
		Test_Case = self.Str_Test_Case_Path.get()
		Serial = self.TextSerial.get().replace('\n','')
		#MyDB = self.Function_Import_DB(self.DB_Path)
		self.Automation_ScriptLoader = Process(target=Function_Generate_Testcase, args=(self.Status_Queue, self.Result_Queue, DB, Test_Case,))
		#self.Data_Compare_Process = Process(target=Old_Function_Compare_Excel, args=(self.Status_Queue, self.Process_Queue, Old_File, New_File, Output_Result, Sheet_Name, Index_Col, self.Background_Color, self.Font_Color,))
		self.Automation_ScriptLoader.start()
		self.after(DELAY1, self.Wait_For_Automation_ScriptLoader)	

	def Wait_For_Automation_ScriptLoader(self):
		if (self.Automation_ScriptLoader.is_alive()):
			try:
				Status = self.Status_Queue.get(0)
				if Status != None:
					self.Debugger.insert("end", "\n\r")
					self.Debugger.insert("end", Status)
					self.Debugger.yview(END)
			except queue.Empty:
				pass	
			self.after(DELAY1, self.Wait_For_Automation_ScriptLoader)
		else:
			try:
				self.TestCase = self.Result_Queue.get(0)
				if self.TestCase != None:
					self.Update_Execute_List()
					self.Automation_ScriptLoader.terminate()
					
					self.Debugger.insert("end", "\n\r")
					self.Debugger.insert("end", 'Testcase config loaded.')
					return
				self.Debugger.insert("end", "\n\r")
				self.Debugger.insert("end", 'Fail to load testcase config.')	
				self.Automation_ScriptLoader.terminate()

			except queue.Empty:
				pass
			

###########################################################################################

	def check_device_connection(self):
		process = subprocess.Popen(ADBPATH + ' devices', stdout=subprocess.PIPE, stderr=None, shell=True)
		device = process.communicate()[0].decode("utf-8") 
		print(device)
		return device
	

###########################################################################################

def Function_Execute_Script(
		Status_Queue, Result_Queue, Serial_Nummber, DB_Path, Test_Case_Path, TestCaseObject = [], Execute_Value = None, **kwargs
):
	All = TestCaseObject	
	Status_Queue.put("Importing test case config")

	#os.system( ADBPATH + ' forward tcp:9889 tcp:9889')

	Dir, Name, Ext = Split_Path(Test_Case_Path)
	Result_Folder_Path = Dir + '\\Result' + '_' + Name + '_' + Function_Get_TimeStamp()
	print('Result_Path', Result_Folder_Path)
	Init_Folder(Result_Folder_Path)
	
	Result_File_Path = Result_Folder_Path + '\\' + Name + '_' + Function_Get_TimeStamp() + Ext
	print('Result_File_Path', Result_File_Path)

	AutoTester = Tester(Status_Queue, Serial_Nummber, DB_Path, Result_Folder_Path)

	Connect_Status = AutoTester.Check_Connectivity()
	if Connect_Status == False:
		Status_Queue.put('Device is not connected.')
		return
	if not os.path.isfile(Test_Case_Path):
		Status_Queue.put("Testcase is not exist")
		return

	if All == None or All 	== []:
		Status_Queue.put('Loading test case')
		All = Function_Import_TestCase(Test_Case_Path)
	
	TestCase = All['Testcase']
	TestInfo = All['Info']
	for detail in TestInfo:
		Status_Queue.put( detail +': '+ str(TestInfo[detail]))
	
	#AutoTester.Count_Object('UI_Inventory')
	Test_Type = TestInfo['Type']
	
	if Test_Type == 'GachaTest':
		Data = Function_Import_Data(Test_Case_Path, TestInfo['StringID'])
		Status_Queue.put('Update Gacha Pool')
		AutoTester.Update_Gacha_Pool(DB_Path, TestInfo['Category'], Data)
	
	elif Test_Type in ['ListAutoTest', 'ListManualTest']:
		Data = Function_Import_Data(Test_Case_Path, TestInfo['StringID'])
		print('Data sheet:', Data)
		Status_Queue.put('Update Execution List')
		AutoTester.Update_Execution_List(Data)
		AutoTester.Update_Execution_Value(Execute_Value)

	else:
		#Default type = General
		pass

	Status_Queue.put('Execute test case')
	Start = time.time()
	if Test_Type == 'ListAutoTest':
		print('List data:', AutoTester.Execution_List)
		for current_execute_value in AutoTester.Execution_List:
			print('Current value:', current_execute_value)
			AutoTester.Update_Execution_Value(current_execute_value)
			Function_Execute_TestCase(TestCase, AutoTester, Test_Case_Path, Result_Folder_Path, Test_Type, Status_Queue)
		Result = True
	else:
		Result = Function_Execute_TestCase(TestCase, AutoTester, Test_Case_Path, Result_Folder_Path, Test_Type, Status_Queue)

	if Result == True:
		Status_Queue.put('Test is completed')
	else:
		Status_Queue.put('Fail to execute the test')	
	
	End = time.time()
	Status_Queue.put('Total testing time: ' + str(int(End-Start)) + " seconds.")	
	#AutoTester = V4Test(Status_Queue, Serial_Nummber, DB_Path)
	#AutoTester.Wait_For_Item('UI_BurgerMenu')

def Function_Generate_Testcase(
		Status_Queue, Result_Queue, DB_Path, Test_Case_Path, **kwargs
):
	if not os.path.isfile(Test_Case_Path):
		Status_Queue.put("Testcase is not exist")
		return
		
	All = Function_Import_TestCase(Test_Case_Path)
	
	TestCase = All['Testcase']
	TestInfo = All['Info']
	print('StringID:',TestInfo['StringID'])
	Data = Function_Import_Data(Test_Case_Path, TestInfo['StringID'])
	All['Data'] = Data
	print('Data:', Data)
	for detail in TestInfo:
		Status_Queue.put( detail +': '+ str(TestInfo[detail]))
	Result_Queue.put(All)
	return

###########################################################################################

def main():
	Process_Queue = Queue()
	Result_Queue = Queue()
	Status_Queue = Queue()
	Debug_Queue = Queue()
	
	MyManager = Manager()
	Default_Manager = MyManager.list()
	
	root = Tk()
	My_Queue = {}
	My_Queue['Process_Queue'] = Process_Queue
	My_Queue['Result_Queue'] = Result_Queue
	My_Queue['Status_Queue'] = Status_Queue
	My_Queue['Debug_Queue'] = Debug_Queue

	My_Manager = {}
	My_Manager['Default_Manager'] = Default_Manager

	Windows = Automation_Execuser(root, Queue = My_Queue, Manager = My_Manager,)
	root.mainloop()  


if __name__ == '__main__':
	if sys.platform.startswith('win'):
		multiprocessing.freeze_support()

	main()
