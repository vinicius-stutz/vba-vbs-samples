Dim message, sapi
MsgBox("Please use the headset and listen to what I have to say...")
message = "This is a simple voice test on your Microsoft Windows."
Set sapi = CreateObject("sapi.spvoice")
sapi.Speak message
