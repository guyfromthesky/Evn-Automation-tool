ADBPATH = 'adb'
import os
import subprocess 
try:
    for port in range(50000,65535):
        port = port
        process = subprocess.Popen('telnet 10.89.100.21:' + str(port), stdout=subprocess.PIPE, stderr=None, shell=True)
        return_message = process.communicate()
        for message in return_message:
            if message != None:
                str_message = message.decode("utf-8") 
                if str_message == port:
                    break

        print('str_message', str_message)
        return_port = str_message.replace('\r\n', "")
        
        if return_port == port:
            print('return_port', return_port)
            break
    print('Current port:', port)

except Exception as e:
    print('Error:', e)