Set WshShell = WScript.CreateObject("WScript.Shell")
WshShell.Run "notepad"
WScript.Sleep 500
For i = 1 To 10
  WshShell.SendKeys "AFK"
  WScript.Sleep 400
  WshShell.SendKeys "{ENTER}"
Next