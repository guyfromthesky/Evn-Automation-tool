import cv2
import os




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

Init_Folder(Screenshot)
Init_Folder(Gacha)
Init_Folder(Result)



FileList = os.listdir(Screenshot)
i=0

UniqueList = []

for FileName in FileList:
	File = Screenshot + FileName
	img = cv2.imread(File)
	#cv2.imshow('Img', img)
	#cv2.waitKey(0)

	Duplicated = False
	for template in UniqueList:
		#template = cv2.imread(LookupItem)
		Found = None
		
		result = cv2.matchTemplate(img, template, cv2.TM_CCOEFF_NORMED)
		(_, maxVal, _, maxLoc) = cv2.minMaxLoc(result)
		print(maxVal)
		if maxVal > 0.75:
			Duplicated = True
			break
	if Duplicated == False:
		i+=1
		cv2.imwrite(Result + 'Companion_' + str(i) + ".png", img)
		UniqueList.append(img)

print('Total unique items:', len(UniqueList))		