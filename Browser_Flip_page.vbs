Dim MyVar, WshShell
Set WshShell = WScript.CreateObject("WScript.Shell")
index = 0
Do While index <= 113
	WshShell.AppActivate "Google Chrome"
	WshShell.SendKeys "{PGDN}"
	WScript.Sleep 3000
	index = index + 1
Loop
