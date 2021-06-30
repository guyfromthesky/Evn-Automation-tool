from ppadb.client import Client as AdbClient
from pyminitouch import MNTDevice

try:
	client = AdbClient(host="127.0.0.1", port=5037)
	devices = client.devices()
	Serrial = []
	for device in devices:
		Serrial.append(device.serial)
	
	TextSerial.set_completion_list(Serrial)
	TextSerial.current(0)
except:	
	TextSerial = ""


device = MNTDevice(TextSerial)

device.tap([(400, 400), (600, 600)])