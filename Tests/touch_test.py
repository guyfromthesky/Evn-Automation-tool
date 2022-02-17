from touch import TouchActionBuilder
import os, sys, subprocess
CWD = os.path.abspath(os.path.dirname(sys.argv[0]))
ADBPATH = '\"' + CWD + '\\adb\\adb.exe' + '\"'
os.popen( ADBPATH + ' kill-server')
os.popen( ADBPATH + ' start-server')
#os.popen( ADBPATH + ' usb')
os.popen( ADBPATH + ' shell chmod 755 /data/local/tmp/touch')
os.popen( ADBPATH + ' shell /data/local/tmp/touch')

#os.popen( ADBPATH + ' forward tcp:50001 tcp:8080')
process = subprocess.Popen(ADBPATH + ' forward tcp:50001 tcp:8080', stdout=subprocess.PIPE, stderr=None, shell=True)
return_message = process.communicate()
for message in return_message:
	if message != None:
		str_message = message.decode("utf-8") 
		print('str_message', str_message)
#os.system('curl http://localhost:50001')

x1 = 10
y1 = 10

x1_1, y1_2 = 500, 500
x2_2, y2_2 = 1000, 500
x3_1, y3_1 = 500, 1000
x4_1, y4_1 = 1000, 1000

points1 = [(100,10), (100, 1000)]
points2 = [(x1_1,y1_2), (x2_2, y2_2)]
points3 = [(x1_1,y1_2), (x2_2, y2_2), (x3_1, y3_1)]
points4 = [(x1_1,y1_2), (x2_2, y2_2), (x3_1, y3_1), (x4_1, y4_1)]

one_second = 1000
half_econd = 500

th = TouchActionBuilder()
#th.tap(x1, y1, one_second).execute_and_reset()
#th.multifinger_tap(points1, 1000).execute_and_reset()
# th.doubletap(10,10).execute_and_reset()
# th.multifinger_doubletap(points2).execute_and_reset()
# th.ntap(x2_2, y2_2, 8, 250).execute_and_reset()
try:
	th.tap(10, 10).execute()
	th.multifinger_ntap(points4, 1, 250).execute_and_reset()
except Exception as e:
	print("Error from tapping", e)
# th.swipe_line(x2_1, y2_2, x2_1, y2_2+200).execute_and_reset()
# th.longpress_and_swipe_line(x2_1, y2_1, x2_2, y2_2).execute_and_reset()
# th.swipe_nline(points3, 50, 5).execute_and_reset()
# th.longpress_swipe_nline(points3).execute_and_reset()