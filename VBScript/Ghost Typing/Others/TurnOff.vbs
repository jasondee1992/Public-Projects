Set WshShell = CreateObject("WScript.Shell")
' Kill'
WshShell.Run "taskkill /im  Cscript.exe", , True
