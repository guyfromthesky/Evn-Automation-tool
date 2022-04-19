from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.actions import interaction

from selenium.webdriver.common.actions.action_builder import ActionBuilder

from selenium.webdriver.common.actions.action_builder import PointerInput
# ...

desired_caps = dict(
    platformName='Android',
    platformVersion='11',
    automationName='uiautomator2',
    deviceName='R58M33SC4XJ',
)

driver = webdriver.Remote('http://localhost:4723/wd/hub', desired_capabilities = desired_caps)

actions = ActionChains(driver)

actions.w3c_actions = ActionBuilder(driver, mouse=PointerInput(interaction.POINTER_TOUCH, 'touch'))
#actions.w3c_actions.pointer_action.move_to_location(x=540, y=940)

actions.w3c_actions.pointer_action.move_to_location(x=340, y=940)
actions.w3c_actions.pointer_action.click_and_hold()
actions.w3c_actions.pointer_action.pause(750)
actions.w3c_actions.pointer_action.move_to_location(x=740, y=940)
actions.w3c_actions.pointer_action.click_and_hold()
actions.w3c_actions.pointer_action.pause(750)
actions.w3c_actions.pointer_action.move_to_location(x=940, y=940)
actions.w3c_actions.pointer_action.click_and_hold()
actions.w3c_actions.pointer_action.pause(750)
actions.w3c_actions.pointer_action.move_to_location(x=940, y=740)
actions.w3c_actions.pointer_action.click_and_hold()
actions.w3c_actions.pointer_action.pause(750)
# actions.w3c_actions.pointer_action.click()
actions.w3c_actions.pointer_action.pointer_up()
actions.w3c_actions.perform()