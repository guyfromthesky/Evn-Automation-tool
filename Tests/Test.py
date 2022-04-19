from appium import webdriver
import time
from appium.webdriver.common.touch_action import TouchAction
from appium.webdriver.common.multi_action import MultiAction
from appium.webdriver.appium_service import AppiumService



'''

appium_service = AppiumService()
appium_service.start(args=['–address', 'http://localhost', '-p', '4723'])
print('Appium has been started.')
'''

# ...

desired_caps = dict(
    platformName='Android',
    platformVersion='11',
    automationName='uiautomator2',
    deviceName='R58M33SC4XJ',
    unicodeKeyboard = True,
)

driver = webdriver.Remote("http://localhost:4723/wd/hub", desired_caps)
size=driver.get_window_size()
print('Device are ready to use.')

a1 = TouchAction()
a1.press(10, 20)
a1.move_to(10, 200)
a1.release()

a2 = TouchAction()
a2.press(10, 10)
a2.move_to(10, 100)
a2.release()

ma = MultiAction(driver)
ma.add(a1, a2)
ma.perform()

