import importlib.machinery
import importlib.util
from libs.automation_driver import *
# Import mymodule


def load_new_function(path):
    loader = importlib.machinery.SourceFileLoader( 'mymodule', path)
    spec = importlib.util.spec_from_loader( 'mymodule', loader )
    mymodule = importlib.util.module_from_spec( spec )
    loader.exec_module( mymodule )

    # Use mymodule
    mymodule.say_hello()

print('Enter lib path:')
text = input()

loader = importlib.machinery.SourceFileLoader( 'mymodule', text)
spec = importlib.util.spec_from_loader( 'mymodule', loader )
mymodule = importlib.util.module_from_spec( spec )
loader.exec_module( mymodule )

#Auto.append_action_list(type = 'Action', name = mymodule.__name__ , argument = {'template_path':'current_area', 'match_rate': 'float', 'timeout': 'int'}, description= '')

print('Enter lib path:')

text = input()

load_new_function(text)




