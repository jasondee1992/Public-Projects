Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad"
WScript.Sleep 500
    For i = 1 To 8640000
        Rtn=WshShell.AppActivate(myPID)
        WshShell.SendKeys "TEST"
        WScript.Sleep 1000
        WshShell.SendKeys "{ENTER}"
    next
