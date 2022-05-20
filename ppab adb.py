import time
from ppadb.client import Client

adb = Client(host='127.0.0.1', port=5037)
device = adb.device('R58M33SC4XJ')


def sendevent(_type, _code, _value, _deviceName='/dev/input/event4'):
    last_time = time.time()
    command = f"su -c 'sendevent {_deviceName} {_type} {_code} {_value}'"
    print(command)
    device.shell(f"su -c 'sendevent {_deviceName} {_type} {_code} {_value}'")
    #print(time.time()-last_time)

def TapScreen(x, y):
    sendevent(EV_ABS, ABS_MT_ID, 0)
    sendevent(EV_ABS, ABS_MT_TRACKING_ID, 0)
    sendevent(1, 330, 1)
    sendevent(1, 325, 1)
    sendevent(EV_ABS, ABS_MT_PRESSURE, 5)
    sendevent(EV_ABS, ABS_MT_POSITION_X, x)
    sendevent(EV_ABS, ABS_MT_POSITION_Y, y)
    sendevent(EV_SYN, SYN_REPORT, 0)
    sendevent(EV_ABS, ABS_MT_TRACKING_ID, -1)
    sendevent(1, 330, 0)
    sendevent(1, 325, 0)
    sendevent(EV_SYN, SYN_REPORT, 0)

EV_ABS             = 3
EV_SYN             = 0
SYN_REPORT         = 0
ABS_MT_ID          = 47
ABS_MT_TOUCH_MAJOR = 48
ABS_MT_POSITION_X  = 53
ABS_MT_POSITION_Y  = 54
ABS_MT_TRACKING_ID = 57
ABS_MT_PRESSURE    = 58

TapScreen(1000, 500)


