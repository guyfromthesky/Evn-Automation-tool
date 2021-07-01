"./adb/adb.exe" start-server
"./adb/adb.exe" push ./libs/arm64-v8a/touch /data/local/tmp

"./adb/adb.exe" shell chmod 755 /data/local/tmp/touch
"./adb/adb.exe" shell /data/local/tmp/touch
"./adb/adb.exe" forward tcp:9889 tcp:9889
cmd /k