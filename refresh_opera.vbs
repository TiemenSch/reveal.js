Set WshShell = WScript.CreateObject("WScript.Shell") 
WshShell.AppActivate("Opera")
WScript.Sleep 500
WshShell.SendKeys "{F5}" 