#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#include <Array.au3>
#region --- Internal functions Au3Recorder Start ---
Global $arr[9999][3]
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
_WinWaitActivate("DuoSubscribe Version 5.0 - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\trisha.xlsx', 0, 0)
$i = 600
While $i < 899
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseClick("left",419,175,1)
Sleep(75)
MouseClick("left",460,248,1)
Send($sCellValue)
Sleep(150)
Send("{ENTER}")
Sleep(5000)
;~ MouseMove(996,564)
;~ MouseDown("left")
;~ MouseMove(910,554)
;~ MouseUp("left")
MouseClick("left",996,564,1)
Send("{CTRLDOWN}a{CTRLUP}")
Send("{CTRLDOWN}c{CTRLUP}")
Sleep(130)
MouseClick("left",960,591,1)
Send("{CTRLDOWN}v{CTRLUP}")
Sleep(120)
MouseClick("left",313,175,1)
Sleep(5000)
Send("{ENTER}")
Sleep(140)
ClipPut("")
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

