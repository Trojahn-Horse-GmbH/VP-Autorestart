Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run """C:\Program Files (x86)\VPDesk\VPDesk.exe """ & " ""-Dserver=https://IhreDomeain/vplanning7"""
Do Until Success = True
    Success = objShell.AppActivate("Visual Planning")
    Wscript.Sleep 1000
Loop

Wscript.Sleep 10000
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "kennwort"
objShell.SendKeys "{ENTER}"