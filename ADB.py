import os, sys
os.system('adb kill-server')
os.system('adb start-server')

CPU = os.popen('adb shell getprop ro.product.cpu.abi').read().replace('\n', "")
cwd = str(os.path.abspath(os.path.dirname(sys.argv[0])))
#cmdln = 'adb push %s/libs/%s/touch /data/local/tmp' % (cwd, CPU)
os.system('adb push %s/libs/%s/touch /data/local/tmp' % (cwd, CPU)) 
#adb push ./libs/x86_64/touch /data/local/tmp

os.system('adb shell chmod 755 /data/local/tmp/touch')
os.system('adb shell /data/local/tmp/touch')






#db shell data/local/tmp/touch
#adb forward tcp:9889 tcp:9889
#curl -d '[{"type":"down", "contact":0, "x": 100, "y": 100, "pressure": 50}, {"type": "commit"}, {"type": "up", "contact": 0}, {"type": "commit"}]' http://localhost:9889