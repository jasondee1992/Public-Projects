Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad"
WScript.Sleep 500
if objProcess = "notepad" then
vFound = True
else:
    For i = 1 To 10
        Rtn=WshShell.AppActivate(myPID)
        WshShell.SendKeys "TEST"
        WScript.Sleep 400
        WshShell.SendKeys "{ENTER}"
    next
End if
