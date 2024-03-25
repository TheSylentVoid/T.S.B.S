set x=createobject("wscript.shell")

x.sendkeys "^"+"{ESC}"
wscript.sleep 1000
x.sendkeys "google chrome"
wscript.sleep 1000
x.sendkeys "{ENTER}"
wscript.sleep 500
x.sendkeys "https://www.youtube.com/watch?v=dQw4w9WgXcQ"
wscript.sleep 1000
x.sendkeys "{ENTER}"
