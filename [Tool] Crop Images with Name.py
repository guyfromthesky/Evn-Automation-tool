import cv2
import os
import numpy as np
import re

try:
    from PIL import Image
except ImportError:
    import Image
import pytesseract
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'


cwd = os.path.dirname(os.path.realpath(__file__))
Screenshot  = cwd + "\\Screenshot\\"
Gacha  = cwd + "\\Gacha\\"
Result  = cwd + "\\Result\\"

def Init_Folder(FolderPath):
	if not os.path.isdir(FolderPath):
		try:
			os.mkdir(FolderPath)
		except OSError:
			print ("Creation of the directory %s failed" % FolderPath)

def Function_Get_TimeStamp():		
	now = datetime.now()
	timestamp = str(int(datetime.timestamp(now)))			
	return timestamp

def Get_Text_From_Image(name_img):

	name_img = cv2.cvtColor(name_img, cv2.COLOR_BGR2GRAY)
	fxv = 2
	fyv = 2
	name_img = cv2.resize(name_img, None, fx=fxv, fy=fyv, interpolation=cv2.INTER_AREA)
	name_img = cv2.resize(name_img, None, fx=fxv, fy=fyv, interpolation=cv2.INTER_CUBIC)
	name_img = cv2.resize(name_img, None, fx=fxv, fy=fyv, interpolation=cv2.INTER_LINEAR)

	
	kernel = np.ones((1, 1), np.uint8)
	name_img = cv2.dilate(name_img, kernel, iterations=1)
	name_img = cv2.erode(name_img, kernel, iterations=1)
	

	cv2.threshold(cv2.GaussianBlur(name_img, (5, 5), 0), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
	cv2.threshold(cv2.bilateralFilter(name_img, 5, 75, 75), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
	cv2.threshold(cv2.medianBlur(name_img, 3), 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
	cv2.adaptiveThreshold(cv2.GaussianBlur(name_img, (5, 5), 0), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)
	cv2.adaptiveThreshold(cv2.bilateralFilter(name_img, 9, 75, 75), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)
	cv2.adaptiveThreshold(cv2.medianBlur(name_img, 3), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY, 31, 2)
	
	
	Text = pytesseract.image_to_string(name_img, lang="eng")
	Text = Text.split('\n')[0]
	Text = re.sub(r'[\\/\:*"<>\|\.%\$\^&£ ]', '', Text)
	#print('Text:', Text)
	#cv2.imshow('image',name_img)
	#cv2.waitKey(0)

	return Text

Init_Folder(Screenshot)
Init_Folder(Gacha)
Init_Folder(Result)

FileList = os.listdir(Gacha)
i=0
for FileName in FileList:
	i+=1
	File = Gacha + '//' + FileName

	img = cv2.imread(File)
	x=494
	y=287
	h=167
	w=167
	
	deltaX = 50
	deltaY = 188
	nx = x - deltaX
	ny = y + deltaY

	nw = 280
	nh = 32
	for j in range(5):
		newX = x + j*311
		nnewX = nx + j*311
		crop_img = img[y:y+h, newX:newX+w]
		
		name_img = img[ny:ny+nh, nnewX:nnewX+nw]

		Companion_Name = Get_Text_From_Image(name_img)

		cv2.imwrite(Screenshot + Companion_Name + ".png", crop_img) 

	x=338
	y=625
	nx = x - deltaX
	ny = y + deltaY

	for j in range(6):
		newX = x + j*311
		nnewX = nx + j*311
		crop_img = img[y:y+h, newX:newX+w]

		name_img = img[ny:ny+nh, nnewX:nnewX+nw]

		Companion_Name = Get_Text_From_Image(name_img)

		cv2.imwrite(Screenshot + Companion_Name + ".png", crop_img) 

print('Done')
	#cv2.imshow("cropped", crop_img)
	#cv2.waitKey(0)
	
	#cropped_image = img.crop(x, y , x+w, y+h)
	#cropped_image.save()
	#cropped_image.show()
	
	#with open( Screenshot + '_' + str(i) + ".png", "wb") as fp:
	#	fp.write(crop_img)	
	