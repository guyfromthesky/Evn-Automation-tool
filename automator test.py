import uiautomator2 as u2
import time
from multiprocessing import Process, freeze_support
import sys
import threading


d = u2.connect('R58M33SC4XJ')

def touch_touch(x, y):
    d = u2.connect('R58M33SC4XJ')
    print('Tap')
    d.long_click(x, y,1)
    print('Release')

def tap_tap(x, y):
   
    print('Tap')
    d.touch.down(x, y)
    time.sleep(1)
    print('Release')
    d.touch.up(x, y)
    print(locals())


def main():
    '''
    Automation_Processor_1 = Process(target = tap_tap, kwargs= {'x':100, 'y':100},)
    Automation_Processor_2 = Process(target = tap_tap, kwargs= {'x':100, 'y':200},)
    Automation_Processor_3 = Process(target = tap_tap, kwargs= {'x':100, 'y':300},)
    Automation_Processor_4 = Process(target = tap_tap, kwargs= {'x':100, 'y':400},)

    Automation_Processor_1.start()
    Automation_Processor_2.start()
    Automation_Processor_3.start()
    Automation_Processor_4.start()
    '''
    '''
    
    
    '''
    x = 200
    try:
        t1 = threading.Thread( target=tap_tap, args = (x, 200))
        t2 = threading.Thread( target=tap_tap, args = (x, 400))
        t3 = threading.Thread( target=tap_tap, args = (x, 600))
        t4 = threading.Thread( target=tap_tap, args = (x, 800))

        t1.start()
        t2.start()
        t3.start()
        t4.start()

        t1.join()
        t2.join()
        t3.join()
        t4.join()
        

    except Exception as e:
        print("Error: unable to start thread")
        print(e)
    print(d.app_current())
    print('Done')

if __name__ == '__main__':
	if sys.platform.startswith('win'):
		freeze_support()

	main()
