from pywinauto.application import Application

title = 'OrgChart.txt - Notepad'
app = Application().connect(title=title, timeout=10)
if app.software_update.exists(timeout=10):
    print('Found')