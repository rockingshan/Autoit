#region ---Au3Recorder
#include <Date.au3>
#region --- Internal functions Au3Recorder Start ---
Local $aTime
Func _Au3RecordSetup()
Opt('WinWaitDelay',100)
Opt('WinDetectHiddenText',1)
Opt('MouseCoordMode',0)
EndFunc

Func _WinWaitActivate($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

_AU3RecordSetup()
_WinWaitActivate("DuoSubscribe Version 5.0 - Google Chrome","")
MouseClick("left",427,253,1)
Sleep(500)
MouseClick("left",409,451,1)
Sleep(500)
MouseClick("left",449,245,1)
Sleep(500)
_WinWaitActivate("Save As","")
Sleep(250)
MouseClick("left",256,46,1)
Sleep(250)
Send("D:\REPORTS")
Sleep(200)
Send("{ENTER}")
MouseClick("left",238,346,1)
Send("{BACKSPACE}")
Send("Subs_All_"&@YEAR&@MON&@MDAY)
MouseClick("left",516,450,1)
#endregion --- Au3Recorder generated code End ---