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
	for j in range(5):
		newX = x + j*311
		crop_img = img[y:y+h, newX:newX+w]
		cv2.imwrite(Screenshot + 'Companion_Row_1_' + str(i) + "_" + str(j)+ ".png", crop_img) 
	
	x=338
	y=625

	for j in range(6):
		newX = x + j*311
		crop_img = img[y:y+h, newX:newX+w]
		cv2.imwrite(Screenshot + 'Companion_Row_2_' + str(i) + "_" + str(j)+ ".png", crop_img) 


	#cv2.imshow("cropped", crop_img)
	#cv2.waitKey(0)
	
	#cropped_image = img.crop(x, y , x+w, y+h)
	#cropped_image.save()
	#cropped_image.show()
	
	#with open( Screenshot + '_' + str(i) + ".png", "wb") as fp:
	#	fp.write(crop_img)	
	