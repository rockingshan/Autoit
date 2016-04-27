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
#endregion
_WinWaitActivate("DuoSubscribe Version 5.0 - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\CABLE.xls', 0, 0)
$i = 2
While $i < 79
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	MouseClick("left",578,212,1)
	Sleep(75)
Send($sCellValue)
Sleep(90)
Send("{ENTER}")
Sleep(1600)
MouseClick("left",517,264,1)
Sleep(120)
Send("{ENTER}")
Sleep(340)
MouseClick("left",616,318,1)
Send("{CTRLDOWN}c{CTRLUP}")
_WinWaitActivate("Untitled - Notepad","")
Send($sCellValue)
Send("{TAB}")
Send("{CTRLDOWN}v{CTRLUP}")
Send("{ENTER}")
Sleep(75)
_WinWaitActivate("DuoSubscribe Version 5.0 - Mozilla Firefox","")
ClipPut("")
MouseClick("left",375,172,1)
#endregion --- Au3Recorder generated code End ---
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


