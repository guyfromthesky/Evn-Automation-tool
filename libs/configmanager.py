#Regular expression handling
import re
import configparser
#http request and parser
# Base64 convert
import base64
#System lib
import os
import sys

class ConfigLoader:
	def __init__(self, section_name):
		self.basePath = os.path.abspath(os.path.dirname(sys.argv[0]))
		if sys.platform.startswith('win'):
			self.appdata = os.environ['APPDATA'] + '\\AutoTester'
		else:
			self.appdata = os.getcwd() + '\\AutoTester'
		self.CWD = os.path.abspath(os.path.dirname(sys.argv[0]))
		# Config file
		self.Auto_Tool_Config_Path = self.appdata + '\\auto_tool.ini'

		# Folder
		self.Config = {}

		# Generate app folder (os.environ['APPDATA'] + '\\Nexon OCR Too)
		self.section_name = section_name
		self.initFolder()

		# Set default value:
		self.Auto_Tool_Init_Setting()
	
		#self.Initconfig()
		#self.Init_DB_Config()
		

	def initFolder(self):
		'''
		Create the config folder incase it's not existed
		'''
		if not os.path.isdir(self.appdata):
			print("No app data folder, create one")
			#No config, create one
			try:
				os.mkdir(self.appdata)
			except OSError:
				print ("Creation of the directory %s failed" % self.appdata)
			else:
				print ("Successfully created the directory %s " % self.appdata)
			#Check local database

	def Auto_Tool_Init_Setting(self):
		'''
		Init the value for OCR TOOL
		'''

		config_path = self.Auto_Tool_Config_Path

		if not os.path.isfile(config_path):
			config = configparser.ConfigParser()
			config.add_section(self.section_name)

			with open(config_path, 'w') as configfile:
				config.write(configfile)

		
		config = configparser.ConfigParser()
		config.read(config_path)
		
		self.Init_Config_Option(config, self.section_name, 'tess_path', '', True)
		self.Init_Config_Option(config, self.section_name, 'tess_data', '', True)
		self.Init_Config_Option(config, self.section_name, 'db_path', self.CWD + '\\DB\\db.csv', True)
		self.Init_Config_Option(config, self.section_name, 'scan_lang', 'eng')
		self.Init_Config_Option(config, self.section_name, 'app_lang', 'en')
		self.Init_Config_Option_Numberic(config, self.section_name, 'resolution', 0)
		self.Init_Config_Option(config, self.section_name, 'scan_type', "Normal")
		


		with open(config_path, 'w') as configfile:
			config.write(configfile)
	
	# Function will load the value from selected option.
	# If value does not exist, return the default value
	def Init_Config_Option(self, Config_Obj, Section, Option, Default_Value, Encoded = False):
		'''
		Set the default config for the application
		Config_Obj: @dict - Config object
		Section: @string - Section name (of the ini structure)
		Option: @string - Option name (of the ini structure)
		Default_Value: @string/int - value of the option(of the ini structure)
		Encoded: @Bool - Base64 encode - Use to encode path config. 
		'''
		# Config does not exist
		if not Section in self.Config:
			self.Config[Section] = {}
		if Encoded == True:
			if isinstance(Default_Value,str):
				if len(Default_Value) > 0:
					print('Default_Value', Default_Value)
					Default_Value =  str(base64.b64encode(Default_Value.encode('utf-8')))
					print('Default_Value', Default_Value)
					Default_Value = re.findall(r'b\'(.+?)\'', Default_Value)[0]
			
		# Config does not have that section
		if not Config_Obj.has_section(Section):
			
			Config_Obj.add_section(Section)
			Config_Obj.set(Section, Option, str(Default_Value))
			self.Config[Section][Option] = Default_Value
		# Config have that section
		else:
			# The section does not have that option

			if not Config_Obj.has_option(Section, Option):
				
				Config_Obj.set(Section, Option, str(Default_Value))
				if Encoded == False:
					self.Config[Section][Option] = Default_Value
				else:
					self.Config[Section][Option] = base64.b64decode(Default_Value).decode('utf-8') 	
			# The section have that option
			else:
				Value = Config_Obj[Section][Option]
				if Encoded == False:
					if Value.isnumeric():
						self.Config[Section][Option] = int(Config_Obj[Section][Option])
					else:	
						self.Config[Section][Option] = Config_Obj[Section][Option]
				else:
					self.Config[Section][Option] = base64.b64decode(Config_Obj[Section][Option]	).decode('utf-8') 	

	def Init_Config_Option_Numberic(self, Config_Obj, Section, Option, Default_Value, Encoded = False):
		# Config does not exist
		if not Section in self.Config:
			self.Config[Section] = {}
		# Config does not have that section
		if not Config_Obj.has_section(Section):
			Config_Obj.add_section(Section)
			Config_Obj.set(Section, Option, str(Default_Value))
			self.Config[Section][Option] = Default_Value
		# Config have that section
		else:
			# The section does not have that option
			if not Config_Obj.has_option(Section, Option):
				Config_Obj.set(Section, Option,str(Default_Value))
				self.Config[Section][Option] = Default_Value
			# The section have that option
			else:
				self.Config[Section][Option] = int(Config_Obj[Section][Option])

	def Config_Save_Path(self, Config_Obj, Section, Path_Value, Default_Value):
		
		
		if not Section in self.Config:
			self.Config[Section] = {}

		if not Config_Obj.has_section(Section):
			Config_Obj.add_section(Section)

		Raw_Encoded_Path =  str(base64.b64encode(Path_Value.encode('utf-8')))
		Encoded_Path = re.findall(r'b\'(.+?)\'', Raw_Encoded_Path)[0]

		Option = 'path'
		if not Config_Obj.has_option(Section, Option):
			Config_Obj.set(Section, Option, str(Default_Value))
			self.Config[Section][Option] = Default_Value
		else:
			Config_Obj[Section][Option] = Encoded_Path
			self.Config[Section][Option] = Encoded_Path
		
	def Config_Load_Path(self, Config_Obj, Section, Default_Value = ''):
		if not Section in self.Config:
			self.Config[Section] = {}

		if not Config_Obj.has_section(Section):
			Config_Obj.add_section(Section)

		Option = 'path'
		if Config_Obj.has_section(Section):
			if Config_Obj.has_option(Section, Option):
				Raw_Path = Config_Obj[Section][Option]
				if Raw_Path != '':
					Path = base64.b64decode(Raw_Path).decode('utf-8')
					self.Config[Section][Option] = Path
				else:
					Config_Obj.set(Section, Option, str(Default_Value))
					self.Config[Section][Option] = Default_Value
			else:
				Config_Obj.set(Section, Option, str(Default_Value))
				self.Config[Section][Option] = Default_Value
		else:
			Config_Obj.set(Section, Option, str(Default_Value))
			self.Config[Section][Option] = Default_Value

	def Get_Config(self, FileName, Section, Option, Default_Value = None, Encode = False):
		'''
		Get the value from the config file
		FileName: @str - Ini path
		Config_Obj: @dict - Config object
		Section: @string - Section name (of the ini structure)
		Option: @string - Option name (of the ini structure)
		Default_Value: @string/int - value of the option(of the ini structure)
		Encoded: @Bool - Base64 encode - Use to encode path config. 
		'''

		if FileName in self:
			config_path = self.FileName
		else:
			return Default_Value

		if not os.path.isfile(config_path):
			return Default_Value

		Config_Obj = configparser.ConfigParser()
		Config_Obj.read(config_path)

		if not Config_Obj.has_section(Section):
			return Default_Value

		if not Config_Obj.has_option(Section, Option):	
			return Default_Value

		Value = Config_Obj[Section][Option]
		if Value != '':
			if Encode == True:
				return base64.b64decode(Value).decode('utf-8')
			else:
				return Value
		else:
			return Default_Value

	def Save_Config(self, Config_Path, Section, Option, Default_Value = None, Encode = False):	
		'''
		Save the target value to the config file
		Config_Obj: @dict - Config object
		Section: @string - Section name (of the ini structure)
		Option: @string - Option name (of the ini structure)
		Default_Value: @string/int - value of the option(of the ini structure)
		Encoded: @Bool - Base64 encode - Use to encode path config. 
		'''

		if Encode == True:
			Default_Value =  str(base64.b64encode(Default_Value.encode('utf-8')))
			Default_Value = re.findall(r'b\'(.+?)\'', Default_Value)[0]

		if not os.path.isfile(Config_Path):
			config = configparser.ConfigParser()
			with open(Config_Path, 'w') as configfile:
				config.write(configfile)

		Config_Obj = configparser.ConfigParser()
		Config_Obj.read(Config_Path)

		if not Config_Obj.has_section(Section):
			Config_Obj.add_section(Section)
			Config_Obj.set(Section, Option, str(Default_Value))

			self.Config[Section] = {}
			self.Config[Section][Option] = Default_Value
		else:
				
			if not Config_Obj.has_option(Section, Option):
				Config_Obj.set(Section, Option, str(Default_Value))
				self.Config[Section][Option] = Default_Value
			else:
				Config_Obj.set(Section, Option, str(Default_Value))

		with open(Config_Path, 'w') as configfile:
			Config_Obj.write(configfile)
		
	def Refresh_Config_Data(self):
		'''
		Reset the config for the application
		'''
		self.Ocr_Tool_Init_Setting()