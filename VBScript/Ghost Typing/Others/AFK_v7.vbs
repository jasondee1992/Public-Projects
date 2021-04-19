Set WshShell = CreateObject("WScript.Shell")
    for i=0 to 1000   'this loop will continue about 30 sec if this not enough increase this number
        Rtn=WshShell.AppActivate(myPID)  'have to be the windows title of application or its process ID
        If Rtn = True Then 
            WshShell.SendKeys "......."                                  ' send key you like
            wscript.sleep 100                ' stop execute next line until finish close app
             End If 
        wscript.sleep 100
    Next