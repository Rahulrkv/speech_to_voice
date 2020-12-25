Dim message, sapi
message=InputBox("text here now")
Set sapi=Createobject("sapi.spvoice")
sapi.Speak message