Dim fso, ws, file, APP_PATH
Set ws = WScript.CreateObject("WScript.Shell")
APP_PATH = """C:\Program Files\Internet Explorer\iexplore.exe"""
ws.run APP_PATH

do
set shl = createobject("wscript.shell")
shl.sendkeys "I eat donkey balls"
wscript.sleep 8000
loop
