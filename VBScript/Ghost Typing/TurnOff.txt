Set WshShell = CreateObject("WScript.Shell")
' Kill'
WshShell.Run "taskkill /F /IM wscript.exe /T", , True
