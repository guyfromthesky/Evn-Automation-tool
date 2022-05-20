from selenium.webdriver.common.touch_action import TouchAction
from selenium import webdriver

# Define desired capabilities
desired_caps = {
    "deviceName": "ac******",
    "platformName": "Android",
    "appPackage": "com.android.dialer",
    "noReset": "true",
    "appActivity": "com.oneplus.contacts.activities.OPDialtactsActivity"
}

driver = webdriver.Remote("http://127.0.0.1:4723/wd/hub", desired_caps)
driver.implicitly_wait(20)


actions = TouchAction(driver)
actions.tap_and_hold(20, 20)
actions.move_to(10, 100)
actions.release()
actions.perform()