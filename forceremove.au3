#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#region --- Internal functions Au3Recorder Start ---
Global $Paused
HotKeySet("{PAUSE}", "TogglePause")
HotKeySet("{ESC}", "Terminate")
HotKeySet("+!d", "ShowMessage") ;Shift-Alt-d
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
#endregion --- Internal functions Au3Recorder End ---

_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\silver remove past.xls', 0, 0)
$i = 103
While $i < 149
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",436,175,1)
Sleep(200)
MouseClick("left",750,254,1)
Sleep(300)
MouseClick("left",591,258,1)
Sleep(400)
MouseClick("left",583,322,1)
Sleep(200)
MouseClick("left",720,253,1)
Sleep(200)
Send($sCellValue)
MouseMove(937,251)
Sleep(200)
MouseDown("left")
MouseMove(937,249)
Sleep(200)
MouseUp("left")
Sleep(200)
MouseClick("left",694,316,1)
Sleep(700)
MouseClick("left",843,591,1)
Sleep(200)
MouseClick("left",485,226,1)
Sleep(200)
MouseClick("left",553,269,1)
Sleep(200)
MouseClick("left",573,390,1)
MouseClick("left",590,211,1)
Sleep(200)
MouseClick("left",346,271,1)
Sleep(2500)
MouseClick("left",639,220,1)
Sleep(400)
MouseClick("left",351,276,1)
Sleep(700)
MouseClick("left",702,221,1)
Sleep(400)
MouseClick("left",892,241,1)
MouseClick("left",899,270,1)
Sleep(2500)
MouseClick("left",352,306,1)
Sleep(400)
MouseClick("left",349,331,1)
Sleep(400)
MouseClick("left",752,222,1)
Sleep(400)
MouseMove(915,261)
Sleep(900)
MouseDown("left")
Sleep(900)
MouseMove(916,261)
MouseUp("left")
Sleep(600)
MouseClick("left",914,272,1)
Sleep(2500)
MouseClick("left",347,302,1)
Sleep(900)
MouseClick("left",349,331,1)
Sleep(400)
MouseClick("left",325,173,1)
Sleep(2200)
MouseMove(677,467)
Sleep(900)
MouseDown("left")
MouseMove(677,468)
MouseUp("left")
MouseClick("left",432,181,1)
Sleep(400)
MouseMove(380,214)
Sleep(400)
MouseDown("left")
Sleep(400)
MouseClick("left",380,214,1)
Sleep(400)
$i = $i +1
   WEnd
#endregion --- Au3Recorder generated code End ---

While 1
    Sleep(100)
WEnd
;;;;;;;;

Func TogglePause()
    $Paused = Not $Paused
    While $Paused
        Sleep(100)
        MsgBox(4096, "", "Script is Paused.")
    WEnd
    ToolTip("")
EndFunc   ;==>TogglePause

Func Terminate()
    Exit 0
EndFunc   ;==>Terminate

Func ShowMessage()
    MsgBox(4096, "", "This is a message.")
EndFunc   ;==>ShowMessage
#endregion --- Au3Recorder generated code End ---
