#include <Date.au3>
#include <Process.au3>
#comments-start
	MsgBox(4096, "", "This won't be executed")
	MsgBox(4096, "", "Or this")
#comments-end
;~ Func _Au3RecordSetup()
;~ Opt('WinWaitDelay',100)
;~ Opt('WinDetectHiddenText',1)
;~ WinWait("[CLASS:Notepad]")Opt('MouseCoordMode',0)
;~ EndFunc
;~ Func _WinWaitActivate($title,$text,$timeout=0)
;~ 	WinWait($title,$text,$timeout)
;~ 	If Not WinActive($title,$text) Then WinActivate($title,$text)
;~ 	WinWaitActive($title,$text,$timeout)
;~ EndFunc

;~ _AU3RecordSetup()

;~ _WinWaitActivate("Opening Report.xlsx","")
;~ MouseClick("left",230,100,1)
;~ MouseClick("left",154,276,1)
;~ MouseClick("left",10,30,1)


;~ DllCall("User32.dll", "int", "LockWorkStation")
;_RunDos("start http://223.29.207.131/duosubscribe5/ui/")
Run("notepad.exe")
WinWait("Untitled - Notepad")
Sleep(1500)
;~ ControlSend("[CLASS:MozillaWindowClass]", "", "GeckoPluginWindow1", "This is a line of text in the notepad window")
ControlClick("[CLASS:Notepad]", "", "Edit1", "left", "1", "16", "33")
;~ ControlSend("DuoSubscribe Version 5.0 - Mozilla Firefox", "", "GeckoPluginWindow1", "Admin")
ControlSend("[CLASS:Notepad]", "", "Edit1", "This is a line of text in the notepad window")

