

from ppadb.client import Client as AdbClient
import os, sys
import cv2
import numpy as np
import time
#import imutils

from libs.general_function import *

#cwd = os.path.dirname(os.path.realpath(__file__))
CWD = os.path.abspath(os.path.dirname(sys.argv[0]))

Default_Screenshot_Folder  = CWD + "\\" "Screenshot"

Init_Folder(Default_Screenshot_Folder)

from openpyxl import load_workbook

# Return result format:
# {
# 	"Type": Execute/Result/..,
#	"Status": True/False,
#	"Details": {},
#	"Screenshot": Array,
# }

class Automation:
	# Serial: Device's serial
	# DB: Database's Path.
	def __init__(self, Status_Queue, Serial = None, DB_Path = None, Result_Folder_Path = None):
		self.Debugger = Status_Queue
		
		self.Client = AdbClient(host="127.0.0.1", port=5037)

		if Serial != None:	
			if self.Client != None:	
				self.Device = self.Client.device(Serial)
			else:
				self.Device = None
		else:
			self.Device = None

		self.Gacha_Pool = []
		self.Execution_List = []
		self.Current_Value = None
		
		self.action_list = []
		
		self.Item_List = []
		self.Result_Array = []
		if DB_Path != None:
			self.DB_Path = self.Get_Folder(DB_Path)
			self.Function_Import_DB()
		
		self.Result_Path = Result_Folder_Path

		

		self.tess_path = None
		self.tess_data = None
		self.tess_lang = None
		self.OCR = False

		self.Update_Action_List()

	def append_action_list(self, type = None, name = None, argument = [], description = ''):
	
		_action = {}
		_action['type'] = type
		_action['name'] = name
		_action['arg'] = argument
		_action['description'] = description

		self.action_list.append(_action)
		return _action

	def Function_Import_TestCase(self, TestCase_Object):
		
		to_eval_list = []
		preflix = 'Controller' + '.'	
		# TestCase_Object = List object
		for _index in range(0, TestCase_Object):
			Step = TestCase_Object[_index]
			Result = None
			Stepname = preflix + Step['Name']
			if Step['Name'] == "Action":
				print('Action step')

				Arg = Step['Argument']
				TempArg = []
				for temp_arg in Arg:
					TempArg.append('\"' + str(temp_arg) + '\"')

				TempArg = str(','.join(TempArg))
				if len(Arg) > 0:
					toEval = Stepname + '(' + str(TempArg) + ')'
				else:
					toEval = Stepname + '()'	
				to_eval_list.append(toEval)

			elif Step['Name'] == "Condition":
				print('Condition step')

			elif Step['Name'] == "Get_Result":
				print('Get Result step')
			
			elif Step['Name'] == "Update_Variable":
				print('Update Variable step')		

			else:
				#Loop
				_loop_amount = 1
				if Step['Action'] == 'Loop':
					# Normal loop
					_loop_amount = Step['Arg'][0]
				_loop_steps = []	
				for i in range(0, _loop_amount):
					_temp_loop_steps = []
					while True:
						_temp_index += 1
						LoopStep = TestCase_Object[_temp_index]
						LStepname = preflix + LoopStep['Name']
						#LArg = NewStep['Argument']
						Arg = LoopStep['Argument']
						TempArg = []
						#for temp_arg in LArg:
						for temp_arg in Arg:
							TempArg.append('\"' + str(temp_arg) + '\"')
						TempArg = str(','.join(TempArg))	
						toEval = LStepname + '(' + TempArg + ')'
						_temp_loop_steps.append(toEval)


	def Function_Import_DB(self):
		#self.StringID = []	
		self.UI = {}
		if self.DB_Path != None:
			if (os.path.isfile(self.DB_Path)):
				xlsx = load_workbook(self.DB_Path, data_only=True)
				for sheet in xlsx:	
					sheetname = sheet.title			
					print('Adding DB from: ', sheet)
					DB_Name = sheetname
					Col_StringID = ""
					Col_String_EN = ""
					Col_String_KR = ""
					Col_Path = ""

					ws = xlsx[sheet.title]

					database = None
					ListCol = {}

					#Get Col Label and Letter
					for row in ws.iter_rows():
						for cell in row:
							if cell.value == "StringID":
								Col = cell.column_letter
								Row_ColID = cell.row
								Col_StringID = Col
								ListCol['StringID'] = Col_StringID

							if Col_StringID != "":
								database = ws
								lastChar = Col_StringID
								
								while True:
									lastChar = chr(ord(lastChar) + 1)
									try:
										ColLabel = database[lastChar + str(Row_ColID)].value
									except:
										break	
									if ColLabel in ["",None] :
										break
									else:
										ListCol[ColLabel] = lastChar

								
						if database!=  None:
							break		
					# Load data 			
					if database != None:
						for i in range(Row_ColID, database.max_row): 
							StringID = database[Col_StringID + str(i+1)].value
							#self.StringID.append(StringID)
							MyEntry = {}
							for Label in ListCol:
								if Label == 'Path':
									Path = database[ListCol[Label] + str(i+1)].value
									try:
										Path = Correct_Path(Path)
									except:
										Path = None	
									if Path != None:
										if os.path.isfile(Path):	
											#MyEntry['Path'] = Path
											#MyEntry['Image'] = read_img(Path)
											MyEntry[Label] = database[ListCol[Label] + str(i+1)].value

									
								else:
									MyEntry[Label] = database[ListCol[Label] + str(i+1)].value
							if 'Path' in MyEntry:
								self.UI[StringID] = MyEntry	

	def Function_Import_Data(self, TestCase_Path, Data_ID):

		Data_ID = Data_ID.lower()
		if TestCase_Path != None:
			if (os.path.isfile(TestCase_Path)):
				xlsx = load_workbook(TestCase_Path, data_only=True)
				#Entry = {}
				Data = []
				for sheet in xlsx:
					
					sheetname = sheet.title.lower()
					print('sheetname:', sheetname)
					if sheetname.find('data_')  > -1:
						DataName = sheetname.replace('data_', '')
						if DataName ==  Data_ID:
							FirstDataRow = None
							EndDataRow = None

							ws = xlsx[sheet.title]
							#Get Col Label and Letter
							for col in ws.iter_cols(min_row=1, max_col=1):
								for cell in col:
									Value = cell.value
									if Value != None:
										Value = Value.lower()
										print('Value:', Value)
										CurrentRow = cell.row
										
										if Value == 'stringid':
											FirstInfoRow = CurrentRow
										else:
											Data.append(Value)

							return Data

			else:
				return([])		
		else:
			return([])	

	def Function_Execute_TestCase( self,TestSteps, Controller, TestCase_Path, Result_Path, Test_Type, Status_Queue):
		
		preflix = 'Controller' + '.'	
		ResultComment = []
		Export_Result = []


		for Step in TestSteps:
			Result = None
			Stepname = preflix + Step['Name']
			if Step['Name'] != "Loop":
				Arg = Step['Argument']
				TempArg = []
				for temp_arg in Arg:
					TempArg.append('\"' + str(temp_arg) + '\"')

				TempArg = str(','.join(TempArg))
				if len(Arg) > 0:
					toEval = Stepname + '(' + str(TempArg) + ')'
				else:
					toEval = Stepname + '()'	
				Status_Queue.put(str("Execute function: " + Stepname))
				Result = eval(toEval)
				
				if Result['Type'] == "Execute":
					TextResult = 'Result' + str(Stepname) + ': ' + str(Result['Type'])
					Step_Result = {}
					Step_Result['Name'] = str(Step['Name'])
					Step_Result['Status'] = Result['Status']
					if 'Details' in Result:
						Step_Result['Details'] = Result['Details'] 
					Export_Result.append(Step_Result)
					#TestResult = {}
					#TestResult['Name'] = str(Step['Name'])
					#ResultArray.append(TestResult)

				elif Result['Type'] == 'Result':
					TextResult = 'Result' + str(LStepname) + ': ' + str(Result['Type'])
					TestResult = {}
					TestResult['Name'] = str(Step['Name'])
					TestResult['Detail'] = Result['Details']
					Controller.Update_Result_Array(TestResult)

				ResultComment.append(TextResult)

				#Status_Queue.put(str(TextResult))

			else:
				Steps = Step['Step']
				LoopAmount = Step['Amount']
				for i in range(LoopAmount):
					for LoopStep in Steps:
						LStepname = preflix + LoopStep['Name']
						#LArg = NewStep['Argument']
						Arg = LoopStep['Argument']
						TempArg = []
						#for temp_arg in LArg:
						for temp_arg in Arg:
							TempArg.append('\"' + str(temp_arg) + '\"')
						TempArg = str(','.join(TempArg))	
						Status_Queue.put(str("Execute function: " + LStepname))
						toEval = LStepname + '(' + TempArg + ')'
						Result = eval(toEval)
						if Result['Type'] == "Execute":
							TextResult = 'Result ' + str(LStepname) + ': ' + str(Result['Status'])
							Step_Result = {}
							Step_Result['Name'] = str(LoopStep['Name'])
							Step_Result['Status'] = Result['Status']
							if 'Details' in Result:
								Step_Result['Details'] = Result['Details'] 
							Export_Result.append(Step_Result)

						elif Result['Type'] == 'Result':
							TextResult = 'Result ' + str(LStepname) + ': ' + str(Result['Status'])
							TestResult = {}
							TestResult['Name'] = str(LoopStep['Name'])
							TestResult['Detail'] = Result['Details']
							Controller.Update_Result_Array(TestResult)
					
						ResultComment.append(TextResult)
						#CurTime = Function_Get_TimeStamp()
						#ResultLine = CurTime + ': ' + TextResult
						#Status_Queue.put(str(TextResult))
			
			#Dir, Name, Ext = Split_Path(TestCase_Path)

			#Result_Path = 
		if Test_Type not in ['ListAutoTest', 'ListManualTest']:
			Print_Result(TestCase_Path, Export_Result, Result_Path)
			#eval("print(Controller.Result_Array)")

		return True

	def Update_Action_List(self):
		self.append_action_list(type = 'Action', name = 'Tap_Item', argument = {'string_id': 'string_id', 'total_attemp': 'int'}, description= '')
		self.append_action_list(type = 'Action', name = 'Tap_Location', argument = {'location': 'point'}, description= '')
		self.append_action_list(type = 'Action', name = 'Tap_Template', argument = {'image_path': 'string', 'total_attemp': 'int'}, description= '')
		self.append_action_list(type = 'Action', name = 'Relative_Tap', argument = {'string_id': 'string_id', 'Delta_X': 'int', 'Delta_Y': 'int'}, description= '')
		self.append_action_list(type = 'Action', name = 'Send_Tab_Key', argument = None, description= '')
		self.append_action_list(type = 'Get_Result', name = 'Count_Object', argument = {'string_id':'string_id'}, description= '')
		self.append_action_list(type = 'Update_Variable', name = 'Update_Gacha_Pool', argument = {'db_path':'string', 'db_sheet_name': 'string', 'gacha_pool_sheet':'string'}, description= '')
		self.append_action_list(type = 'Update_Variable', name = 'Update_Execution_List', argument = {'Execute_List':'string'}, description= '')
		self.append_action_list(type = 'Update_Variable', name = 'Update_Execution_Value', argument = {'Execute_List':'string'}, description= '')
		self.append_action_list(type = 'Get_Result', name = 'Analyse_Gacha_Acquired', argument = {'total_item_in_gacha': 'int'}, description= '')
		self.append_action_list(type = 'Update_Variable', name = 'Analyse_Gacha_Result', argument = {'total_item_in_gacha': 'int'}, description= '')
		self.append_action_list(type = 'Action', name = 'wait_for_item', argument = {'string_id':'string', 'match_rate': 'float', 'timeout': 'int'}, description= '')

		self.append_action_list(type = 'Action', name = 'Swipe_by_StringID', argument = {'location_A': 'point', 'location_B':'point'}	, description= '')
		self.append_action_list(type = 'Action', name = 'Swipe_by_StringID', argument = {'string_id_A': 'string_id', 'string_id_B':'string_id'}	, description= '')
		self.append_action_list(type = 'Action', name = 'Send_Enter_Key', argument = None, description= '')
		self.append_action_list(type = 'Action', name = 'Input_Text', argument = {'input_text':'string'}, description= '')
		self.append_action_list(type = 'Action', name = 'Input_Current_Value', argument = None, description= '')
		self.append_action_list(type = 'Action', name = 'Tap_Current_Item', argument = None, description= '')
		self.append_action_list(type = 'Action', name = 'Wait_For_Current_Item', argument = None, description= '')
		self.append_action_list(type = 'Action', name = 'Get_Screenshot', argument = None, description= '')
		#self.append_action_list(type = 'Loop', name = 'List_Loop', argument = {'list_name': 'string'}, description= '')
		self.append_action_list(type = 'Loop', name = 'Loop', argument = {'amount': 'int'}, description= '')
		self.append_action_list(type = 'Condition', name = 'If', argument = {'condition': 'string'}, description= '')
		
		if self.OCR == True:
			self.append_action_list(type = 'Action', name = 'Scan_Text', argument = {'scan_area': 'area'}, description= '')
			#self.append_action_list(type = 'Action', name = 'Test_Scan_Text', argument = {'scan_area': 'area', '2nd_scan_area': 'area'}, description= '')
		
		self.append_action_list(type = 'Action', name = 'Tap', argument = {'touch_point': 'point'}, description= '')
		#self.append_action_list(type = 'Action', name = 'Test_Tap', argument = {'touch_point': 'point', '2nd_touch_point': 'point'}, description= '')

	def Get_Current_Screenshot(self):
		Img_Screenshot = self.Device.screencap()
		return Img_Screenshot

	def Get_Folder(self, Path):

		return os.path.dirname(Path)
		 
	def Update_Result_Path(self, Path):
		self.Result_Path = Path

	def Update_DB_Path(self, Path):
		self.DB_Path = Path
		self.Function_Import_DB()

	def Update_Tesseract(self, tess_path, tess_data, tess_lang):
		self.tess_path = tess_path
		self.tess_data = tess_data
		self.tess_lang = tess_lang
		self.OCR = True
		self.Update_Action_List()

	def Update_Serial_Number(self, Serial):
		print('Connect to device:', Serial)
		self.Device = self.Client.device(Serial)


	def Check_Connectivity(self):
		try:
			return self.Device.serial
		except:
			return False	

	def Generate_Result(self, Type = None, Status = None, Details = None, Screenshot = []):
		ReturnResult = {}
		if Type != None:
			ReturnResult['Type'] = Type
		else:
			ReturnResult['Type'] = 'Execute'

		if Status != None:
			if Status == True:
				ReturnResult['Status'] = "Pass"
			else:
				ReturnResult['Status'] = "Fail"	
		else:
			ReturnResult['Status'] = "Fail"
		
		if Details != None:
			ReturnResult['Details'] = Details

		if len(Screenshot) != 0:
			ReturnResult['Screenshot'] = Screenshot

		return 	ReturnResult

	def Update_Result_Array(self, Result):
		self.Result_Array.append(Result)

	def Sleep(self, Time):
		STime = int(Time)/ 1000
		time.sleep(STime)
		return self.Generate_Result(Status = True)

	def Tap(self, touch_point):
		tap(self.Device, touch_point['x'], touch_point['y'])		
		return self.Generate_Result(Status = True)

	def Tap_Item(self, StringID, total_attemp = 5):

		Img_Template = self.UI[StringID]['Image']	
		for i in range(total_attemp):
				
			Img_Screenshot = self.Device.screencap()
			Img_Screenshot = np.asarray(Img_Screenshot)
			Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)
			try:
				result = Get_Item(Img_Screenshot, Img_Template, 0.5)
			except Exception as e:
				result = False
				print('Error from Tap_Item:', e)
			if result:
				tap_object(self.Device, result)
				return self.Generate_Result(Status = True)
 
		
		return self.Generate_Result(Status = False)	

	def Tap_Location(self, location):
		tap_object(self.Device, location)
		return self.Generate_Result(Status = True)
 

	def Tap_Template(self, image_path, total_attemp = 5):
		if not os.path.isfile(image_path):
			return self.Generate_Result(Status = False)

		Img_Template = read_img(image_path)
		
		for i in range(total_attemp):
				
			Img_Screenshot = self.Device.screencap()
			Img_Screenshot = np.asarray(Img_Screenshot)
			Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)
			try:
				result = Get_Item(Img_Screenshot, Img_Template, 0.5)
			except Exception as e:
				result = False
				print('Error from Tap_Item:', e)
			if result:
				tap_object(self.Device, result)
				return self.Generate_Result(Status = True)

		
		return self.Generate_Result(Status = False)

	def Relative_Tap(self, StringID, Delta_X=0, Delta_Y=0):
		Img_Template = self.UI[StringID]['Image']	

		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Loc = Get_Item(Img_Screenshot, Img_Template, 0.5)
		
		Loc['x'] += int(Delta_X)
		Loc['y'] += int(Delta_Y)

		if Loc:		
			tap(self.Device, Loc)
			ResultStatus = True
		else:
			ResultStatus = False
		
		return self.Generate_Result(Status = ResultStatus)	

	#def Four_Touch(self):
	#	Four_Touch()

	#def Three_Touch(self):
	#	Three_Touch()

	def Count_Object(self, StringID):
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)
		Img_Template = self.UI[StringID]['Image']
		ResultStatus, Img = Count_Object(Img_Screenshot, Img_Template)
		return self.Generate_Result(Status = ResultStatus, Screenshot = Img)

	def Scan_Text(self, scan_area):
		try:
			_img = self.Device.screencap()
			imCrop = _img[int(scan_area[1]):int(scan_area[1]+scan_area[2]), int(scan_area[0]):int(scan_area[0]+scan_area[3])]
			text = get_text_from_image(self.tess_path, self.tess_lang, self.tess_data, imCrop)
		except:
			return self.Generate_Result(Status = False)
		return self.Generate_Result(Status = text)

	def Update_Gacha_Pool(self, DB_Path, DB_Sheet, Gacha_Pool):
		self.Gacha_Pool = Function_Import_DB(DB_Path, [DB_Sheet], Gacha_Pool)
		
		return self.Generate_Result(Status = True)

	def Update_Execution_List(self, Execute_List):
		self.Execution_List = Execute_List
		return self.Generate_Result(Status = True)

	def Update_Execution_Value(self, Execute_Value):
		self.Execution_Value = Execute_Value
		return self.Generate_Result(Status = True)

	def Analyse_Gacha_Acquired(self, Gacha_Amount = 11):
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)
		#Img_Screenshot = cv2.cvtColor(Img_Screenshot, cv2.COLOR_GRAY2BGR)
		try:
			Img_Screenshot = resize(Img_Screenshot, 50)
		except:
			print('Fail to Resize')
		
		Gacha_Pool = self.Gacha_Pool
		Gacha_Result = {}

		for StringID in Gacha_Pool:
			if  'Image' in Gacha_Pool[StringID]:
				result = False
				result, Img_Screenshot = Count_Object(Img_Screenshot, Gacha_Pool[StringID]['Image'], 0.95)
				
				if result != False:
					Gacha_Result[Gacha_Pool[StringID]['StringID']] = result
					
		Acquired = 0
		Details = {}
		Details['Data'] = Gacha_Result
		Details['Image'] = Img_Screenshot
		for item in Gacha_Result:
			Acquired += int(Gacha_Result[item])
			
		if 	Acquired != Gacha_Amount:
			Status_Result = False
			Name = Correct_Path('Gacha result_' + Function_Get_TimeStamp() + '.png', 'Test Result')
			cv2.imwrite(Name, Img_Screenshot) 
		else:
			Status_Result = True	
		

		return self.Generate_Result(Type = 'Result', Status = Status_Result, Screenshot = Img_Screenshot, Details = Gacha_Result)

	def Analyse_Gacha_Result(self, Total_Item_In_Gacha= 11):
		Count = 0
		Gacha_Pool = self.Gacha_Pool
		Gacha_Result = {}

		Item_Not_In_Pool = False

		for StringID in self.Gacha_Pool:
			Gacha_Result[StringID] = 0
		
		for Result in self.Result_Array:
			if Result['Name'] == "Analyse_Gacha_Acquired":
				Count+=1
				Within_Count = 0
				for StringID in Result['Detail']:
					ItemAmount = Result['Detail'][StringID]
					Gacha_Result[StringID] += ItemAmount
					Within_Count += ItemAmount

			if Within_Count != Total_Item_In_Gacha:
				Item_Not_In_Pool = True
			
		Fail_Details = ""
		Item_Missing = []
		for StringID in Gacha_Result:
			print(StringID)
			if Gacha_Result[StringID] == 0:
				Item_Missing.append(StringID)
		
		if len(Item_Missing) >0:
			Status_Result = False
			Fail_Details+= "Item didn't acquire: " + str(Item_Missing) + '.\r\n'
		else:
			Status_Result = True	

		if Item_Not_In_Pool:
			Status_Result = False
			Fail_Details+= "Acquired amount not match.\r\n"

		return self.Generate_Result(Type = 'Execute', Status = Status_Result, Details = Fail_Details)

	def wait_for_item(self, StringID, Match_Rate = 0.5, timeout=15):
		Start = time.time()
		Wait_Time = timeout * 1000
		Now = Start
		while (Now - Start) < Wait_Time:
			Now = time.time()
			Img_Screenshot = self.Device.screencap()
			Img_Screenshot = np.asarray(Img_Screenshot)
			Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)
			result, Img_Screenshot = Count_Object(Img_Screenshot, self.UI[StringID]['Image'], Match_Rate)
			if result != False:
				ResultStatus = True
			else:
				ResultStatus = False

		return self.Generate_Result(Status = ResultStatus, Screenshot = Img_Screenshot)

	def Swipe(self, point_A, point_B):
		
		result = swipe(self.Device, point_A, point_B)

		return self.Generate_Result(Status = result)

	def Swipe_by_StringID(self, StringID_A, StringID_B):
		
		#StringID = 'UI_MountsShop'
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Template_A_Path = self.UI[StringID_A]['Path']
		Img_Template_A = read_img(Template_A_Path)
		Loc_A = Get_Item(Img_Screenshot, Img_Template_A)
		
		Template_B_Path = self.UI[StringID_B]['Path']
		Img_Template_B = read_img(Template_B_Path)
		Loc_B= Get_Item(Img_Screenshot, Img_Template_B)

		result = swipe(self.Device, Loc_A, Loc_B)

		return self.Generate_Result(Status = result)

	def Send_Enter_Key(self):
		send_key(self.Device, '66')
		ResultStatus = True

		return self.Generate_Result(Status = ResultStatus)


	def Send_Tab_Key(self):
		send_key(self.Device, '61')
		ResultStatus = True

	

		return self.Generate_Result(Status = ResultStatus)

	def Send_BackKey_Key(self):
		send_key(self.Device, '4')
		ResultStatus = True

		self.append_action_list(type = 'Action', name = 'Send_BackKey_Key', argument = [], description= '')

		return self.Generate_Result(Status = ResultStatus)
		

	def Input_Text(self, Text):
		send_text(self.Device, Text)
		ResultStatus = True

		

		return self.Generate_Result(Status = ResultStatus)

	def Input_Current_Value(self):
		send_text(self.Device, self.Execution_Value)
		ResultStatus = True

	

		return self.Generate_Result(Status = ResultStatus)

	def Tap_Current_Item(self):
		self.tap_item(self.Execution_Value)
		ResultStatus = True

		
		
		return self.Generate_Result(Status = ResultStatus)

	def Wait_For_Current_Item(self):
		self.wait_for_item(self.Execution_Value)	
		ResultStatus = True

		
		return self.Generate_Result(Status = ResultStatus)
	
	def Get_Screenshot(self, Name = 'Screenshot_'):
		Image = self.Device.screencap()
		Img_Name = Correct_Path(Name + Function_Get_TimeStamp() + '.png', self.Result_Path)
		ResultStatus = True
		try:
			with open(Img_Name, "wb") as fp:
				fp.write(Image)
		except:
			ResultStatus = False
		
		return self.Generate_Result(Status = ResultStatus)


