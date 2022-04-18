from appium import webdriver
from appium.webdriver.common.touch_action import TouchAction
from appium.webdriver.common.multi_action import MultiAction
# ...

desired_caps = dict(
    platformName='Android',
    platformVersion='11',
    automationName='uiautomator2',
    deviceName='R58M33SC4XJ',
    app=r'C:\Users\evan\OneDrive - NEXON COMPANY\[Demostration] V4 Gacha test\selendroid-test-app-0.17.0.apk'
)

driver = webdriver.Remote('http://localhost:4723/wd/hub', desired_caps)
'''
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

'''

ma = MultiAction(driver)
ma.add(a1, a2)
ma.perform()