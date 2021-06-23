

from ppadb.client import Client as AdbClient
import os, sys
import cv2
import numpy as np
from datetime import datetime
import time
import imutils

from nxautomation import *

#cwd = os.path.dirname(os.path.realpath(__file__))
cwd = os.path.abspath(os.path.dirname(sys.argv[0]))

Screenshot  = cwd + "\\"

if not os.path.isdir(Screenshot):
	try:
		os.mkdir(Screenshot)
	except OSError:
		print ("Creation of the directory %s failed" % Screenshot)

from openpyxl import load_workbook, worksheet, Workbook


class CallCheat:
	# Serial: Device's serial
	# DB: Database's Path.
	def __init__(self, Serial, DB):
		self.UI = Function_Import_DB(DB)
		self.Project = "V4"
		self.Client = AdbClient(host="127.0.0.1", port=5037)
		for d in adb.devices():
			print(d.serial)
		self.Device = client.device(Serial)

	def Enable_MH(self):
		pass

	def Enable_Console(self):
		pass

	def Count_Object(self, Object_Image_Path):
		
		return

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

	def Tap_Item(self, StringID, ):

		Img_Template = self.UI[StringID]['Image']	

		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		result = Fast_Search(Img_Screenshot, Img_Template, 0.5)
		if result:		
			Tap(self.Device, result)
			ResultStatus = True
		else:
			ResultStatus = False

		return self.Generate_Result(Status = ResultStatus)

	def Relative_Tap(self, StringID, Delta_X=0, Delta_Y=0):
		Img_Template = self.UI[StringID]['Image']	

		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Loc = Fast_Search(Img_Screenshot, Img_Template, 0.5)
		
		Loc['x'] += int(Delta_X)
		Loc['y'] += int(Delta_Y)

		if result:		
			Tap(self.Device, Loc)
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

	def Find_Gacha_Frame(self):
		#Img_Template = self.DB[StringID]['Image']
		image_path = cwd + '\DB\\UI\\Gacha.jpg'
		image = cv2.imread(image_path)
		image = HD_Resize(image) 

		# Grayscale 
		gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY) 
		blur = cv2.medianBlur(gray, 1)
		sharpen_kernel = np.array([[-1,-1,-1], [-1,9,-1], [-1,-1,-1]])
		sharpen = cv2.filter2D(blur, -1, sharpen_kernel)

		# Find Canny edges 
		sharpen = cv2.Canny(gray, 10, 200) 
		
		# Finding Contours 
		# Use a copy of the image e.g. edged.copy() 
		# since findContours alters the image 
		contours, hierarchy = cv2.findContours(sharpen,  
			cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_NONE) 
		
		cv2.imshow('Canny Edges After Contouring', sharpen) 
		cv2.waitKey(0) 
		

		cnts = cv2.findContours(sharpen, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
		cnts = cnts[0] if len(cnts) == 2 else cnts[1]

		min_area = 75
		max_area = 200
		image_number = 0
		for c in cnts:
			area = cv2.contourArea(c)
			if area > min_area and area < max_area:
				x,y,w,h = cv2.boundingRect(c)
				ROI = image[y:y+h, x:x+h]
				print('ROI_{}.png'.format(image_number))
				cv2.imwrite('ROI_{}.png'.format(image_number), ROI)
				cv2.rectangle(image, (x, y), (x + w, y + h), (36,255,12), 2)
				image_number += 1
		
		# Draw all contours 
		# -1 signifies drawing all contours 
		#cv2.drawContours(image, contours, -1, (0, 255, 0), 3) 
		
		cv2.imshow('Contours', image) 
		cv2.waitKey(0) 
		cv2.destroyAllWindows() 

	def Quare_Detection(self):
		filter = False
		#Img_Template = self.DB[StringID]['Image']
		image_path = cwd + '\DB\\UI\\Gacha 2.jpg'
		image = cv2.imread(image_path)
		image = HD_Resize(image) 

		gray = cv2.cvtColor(image,cv2.COLOR_BGR2GRAY)
		edges = cv2.Canny(gray,90,150,apertureSize = 3)
		kernel = np.ones((3,3),np.uint8)
		edges = cv2.dilate(edges,kernel,iterations = 1)
		kernel = np.ones((5,5),np.uint8)
		edges = cv2.erode(edges,kernel,iterations = 1)


		lines = cv2.HoughLines(edges,1,np.pi/180,150)

		cv2.imshow('hough.jpg',edges)
		cv2.waitKey(0) 

		if not lines.any():
			print('No lines were found')
			exit()

		if filter:
			rho_threshold = 15
			theta_threshold = 0.1

			# how many lines are similar to a given one
			similar_lines = {i : [] for i in range(len(lines))}
			for i in range(len(lines)):
				for j in range(len(lines)):
					if i == j:
						continue

					rho_i,theta_i = lines[i][0]
					rho_j,theta_j = lines[j][0]
					if abs(rho_i - rho_j) < rho_threshold and abs(theta_i - theta_j) < theta_threshold:
						similar_lines[i].append(j)

			# ordering the INDECES of the lines by how many are similar to them
			indices = [i for i in range(len(lines))]
			indices.sort(key=lambda x : len(similar_lines[x]))

			# line flags is the base for the filtering
			line_flags = len(lines)*[True]
			for i in range(len(lines) - 1):
				if not line_flags[indices[i]]: # if we already disregarded the ith element in the ordered list then we don't care (we will not delete anything based on it and we will never reconsider using this line again)
					continue

				for j in range(i + 1, len(lines)): # we are only considering those elements that had less similar line
					if not line_flags[indices[j]]: # and only if we have not disregarded them already
						continue

					rho_i,theta_i = lines[indices[i]][0]
					rho_j,theta_j = lines[indices[j]][0]
					if abs(rho_i - rho_j) < rho_threshold and abs(theta_i - theta_j) < theta_threshold:
						line_flags[indices[j]] = False # if it is similar and have not been disregarded yet then drop it now

		filtered_lines = []

		if filter:
			for i in range(len(lines)): # filtering
				if line_flags[i]:
					filtered_lines.append(lines[i])
		else:
			filtered_lines = lines

		X_Line = []
		Y_Line = []

		for line in filtered_lines:
			rho,theta = line[0]
			a = np.cos(theta)
			b = np.sin(theta)
			x0 = a*rho
			y0 = b*rho
			x1 = int(x0 + 1000*(-b))
			y1 = int(y0 + 1000*(a))
			x2 = int(x0 - 1000*(-b))
			y2 = int(y0 - 1000*(a))


			if abs(x1-x2) <= 10:
				X_Line.append(x1)
				#cv2.line(image,(x1,y1),(x2,y2),(0,0,255),2)
			elif abs(y1-y2) <= 10:
				y = int(0.5 * (y1+y2))
				if y >= 100 and y <= 500:
					Y_Line.append(y)
				#cv2.line(image,(x1,y1),(x2,y2),(0,0,255),2)	

		 

		Y_Line.sort()
		Current_Y = 0
		for y in Y_Line:
			if (y - Current_Y) >= 20:
				Current_Y = y
			else:
				Y_Line.remove(y)

		for x in X_Line:
			if (x - Current_X) >= 20:
				Current_X = x
			else:
				X_Line.remove(x)		

		print('Total x lines:', len(X_Line))
		print('Total y lines:', len(Y_Line))
		for y in Y_Line:
			print('y:', y)
			cv2.line(image,(0,y),(1280,y),(0,0,255),2)
		for x in X_Line:
			print('x:', x)
			cv2.line(image,(x,0),(x,720),(0,0,255),2)
			#cv2.line(image,(x1,y1),(x2,y2),(0,0,255),2)

		cv2.imshow('hough.jpg',image)
		cv2.waitKey(0) 
		#cv2.imwrite('hough.jpg',img)


	def Swipe_Down_V4Shop(self):
		StringID = 'UI_MountsShop'
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Img_Template = self.UI[StringID]['Image']
		#Search_All_Object(Img_Screenshot, Img_Template)
		#result = Search_Best_Match(Img_Screenshot, Img_Template)
		result = Fast_Search(Img_Screenshot, Img_Template)
		if result:
			Swipe_Up(self.Device, result, 500)
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
		Img_Template = cv2.imread(Template_Path)
		result = Fast_Search(Img_Screenshot, Img_Template)
		if result:
			Swipe_Up(self.Device, result, -500)
			ResultStatus = True
		else:
			ResultStatus = False

		return self.Generate_Result(Status = ResultStatus)

	def Update_Gacha_Pool(self, DB_Path, DB_Sheet, Pool):
		self.Gacha_Pool = Function_Import_DB(DB_Path, [DB_Sheet], Pool)
		return self.Generate_Result(Status = True)

	def Update_Execution_List(self, DB_Path, DB_Sheet, Pool):
		self.Execution_List = Function_Import_DB(DB_Path, [DB_Sheet], Pool)
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
			Img_Screenshot = Resize(Img_Screenshot, 50)
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

	def Wait_For_Item(self, StringID, Match_Rate = 0.5, timeout=15):
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


	def Get_Gacha_Image_By_Pos(Pos):

		return Gacha_Image

	def Swipe_by_StringID(self, StringID_A, StringID_B):
		StringID = 'UI_MountsShop'
		Img_Screenshot = self.Device.screencap()
		Img_Screenshot = np.asarray(Img_Screenshot)
		Img_Screenshot = cv2.imdecode(Img_Screenshot, cv2.IMREAD_COLOR)

		Template_A_Path = self.UI[StringID_A]['Path']
		Img_Template_A = cv2.imread(Template_A_Path)
		Loc_A = Fast_Search(Img_Screenshot, Img_Template_A)
		
		Template_B_Path = self.UI[StringID_B]['Path']
		Img_Template_B = cv2.imread(Template_B_Path)
		Loc_B= Fast_Search(Img_Screenshot, Img_Template_B)

		if result:
			Swipe(self.Device, Loc_A, Loc_B)
			ResultStatus = True
		else:
			ResultStatus = False

		return self.Generate_Result(Status = ResultStatus)

	def Send_Enter_Key(self):
		Send_Key(self.Device, '66')
		ResultStatus = True
		return self.Generate_Result(Status = ResultStatus)


	def Send_Tab_Key(self):
		Send_Key(self.Device, '61')
		ResultStatus = True
		return self.Generate_Result(Status = ResultStatus)

	def Send_BackKey_Key(self):
		Send_Key(self.Device, '4')
		ResultStatus = True
		return self.Generate_Result(Status = ResultStatus)
		

	def Input_Text(self, Text):
		Send_Text(self.Device, Text)
		ResultStatus = True
		return self.Generate_Result(Status = ResultStatus)

	def Input_Current_Value(self):
		Send_Text(self.Device, self.Execution_Value)
		ResultStatus = True
		return self.Generate_Result(Status = ResultStatus)

	def Tap_Current_Item(self):
		self.Tap_Item(self.Execution_Value)

	def Wait_For_Current_Item(self):
		self.Wait_For_Item(self.Execution_Value)	



#V4 = V4Test()