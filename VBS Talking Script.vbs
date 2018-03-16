Dim message,sapi
message=InputBox("Type what you want to be said")
set sapi = CreateObject("sapi.spvoice")
sapi.Speak message