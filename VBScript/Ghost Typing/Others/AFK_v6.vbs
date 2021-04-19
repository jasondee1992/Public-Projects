Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad"
WScript.Sleep 500
    For i = 1 To 1000
        Rtn=WshShell.AppActivate(myPID)
        WshShell.SendKeys "I'M AFK FOR 10 Minutes."
        WScript.Sleep 400
        WshShell.SendKeys "{ENTER}"
    next
