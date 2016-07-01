Dim message, sapi
Dim ArgObj 
Set ArgObj = WScript.Arguments 

'First parameter
var1 = ArgObj(0) 


message=var1
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak message