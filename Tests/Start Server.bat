adb start-server
adb push ./libs/arm64-v8a/touch /data/local/tmp
adb shell chmod 755 /data/local/tmp/touch
adb shell /data/local/tmp/touch
adb forward tcp:50001 tcp:50001

cmd /k