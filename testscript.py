

from ppadb.client import Client as AdbClient
import os, sys
import cv2
import numpy as np
from datetime import datetime
import time
import imutils

from nxautomation import *

#cwd = os.path.dirname(os.path.realpath(__file__))
CWD = os.path.abspath(os.path.dirname(sys.argv[0]))

Screenshot  = CWD + "\\"

if not os.path.isdir(Screenshot):
	try:
		os.mkdir(Screenshot)
	except OSError:
		print ("Creation of the directory %s failed" % Screenshot)

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
	def __init__(self, Status_Queue, Serial, DB_Path):
		self.Debugger = Status_Queue
		#self.Companion = Function_Import_DB(DB_Path, ['Companion'])
		#self.Project = "V4"
		self.Client = AdbClient(host="127.0.0.1", port=5037)
		self.Device = self.Client.device(Serial)
		self.Gacha_Pool = []
		self.Execution_List = []
		self.Current_Value = None
		self.Item_List = []
		self.Result_Array = []
		self.DB_Path = self.Get_Folder(DB_Path)
		self.UI = Function_Import_DB(DB_Path, ['UI'])


	def Get_Folder(self, Path):

		return os.path.dirname(Path)
		 
	def Add_DB_Path(self, Path):
		return self.DB_Path + "//" + Path

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

	def Four_Touch(self):
		Four_Touch()

	def Three_Touch(self):
		Three_Touch()

	def Nav_V4Shop(self):
		action_result = self.Tap_Item('UI_V4Shop')

		return action_result
		

	def Nav_Exit(self):
		action_result = self.Tap_Item('UI_Exit')
		return action_result


	def Nav_BurgerMenu(self):
		action_result = self.Tap_Item('UI_BurgerMenu')

		return action_result

	def Nav_Inventory(self):
		action_result = self.Tap_Item('UI_Inventory')
		
		return action_result

	def Nav_CompanionsShop(self):
		action_result = self.Tap_Item('UI_CompanionsShop')

		return action_result

	def Tap_If_Exist(self):
		return


	def Count_Object(self, StringID):
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)
		Img_Template = self.UI[StringID]['Image']
		ResultStatus, Img = Count_Object(Img_Screenshot, Img_Template)
		return self.Generate_Result(Status = ResultStatus, Screenshot = Img)


	def Swipe_Down_V4Shop(self):
		StringID = 'UI_MountsShop'
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Img_Template = self.UI[StringID]['Image']
		#Search_All_Object(Img_Screenshot, Img_Template)
		#result = Search_Best_Match(Img_Screenshot, Img_Template)
		result = Get_Item(Img_Screenshot, Img_Template)
		if result:
			swipe_up(self.Device, result, 500)
			ResultStatus = True
		else:
			ResultStatus = False

		return self.Generate_Result(Status = ResultStatus)

	def Swipe_Up_V4Shop(self):
		StringID = 'UI_MountsShop'
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Template_Path = self.UI[StringID]['Path']
		Img_Template = read_img(Template_Path)
		result = Get_Item(Img_Screenshot, Img_Template)
		if result:
			swipe_up(self.Device, result, -500)
			ResultStatus = True
		else:
			ResultStatus = False

		return self.Generate_Result(Status = ResultStatus)

	def Update_Gacha_Pool(self, DB_Path, DB_Sheet, Pool):
		self.Gacha_Pool = Function_Import_DB(DB_Path, [DB_Sheet], Pool)
		return self.Generate_Result(Status = True)

	def Update_Execution_List(self, Data):
		self.Execution_List = Data
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
			Img_Screenshot = _resize(Img_Screenshot, 50)
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

	def Wait_For_Current_Item(self):
		self.wait_for_item(self.Execution_Value)	



#V4 = V4Test()