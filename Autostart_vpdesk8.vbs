Set objShell = WScript.CreateObject("WScript.Shell")
On Error Resume Next
objShell.Run """taskkill /F /IM javaw.exe"""
objShell.Run """<Path_to>\VPDeskLauncher.exe""" & " ""-Dserver=<URL_Ihres_Visual_Plannings>"""
Wscript.Sleep 5000    
Do Until Success = True
    Success = objShell.AppActivate("VPDesk 8")
    Wscript.Sleep 1000
Loop

Wscript.Sleep 10000
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "{TAB}"
objShell.SendKeys "<PASSWORD>"
objShell.SendKeys "{ENTER}"
