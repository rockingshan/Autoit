#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#region --- Internal functions Au3Recorder Start ---
Global $arr[9999][4]
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
;;;80% zoom for modify customer
_AU3RecordSetup()
#endregion --- Internal functions Au3Recorder End ---
;Local $oExcel2 = _ExcelBookNew()
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\card.xlsx', 0, 0)
$i = 500
While $i < 600
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	_WinWaitActivate("safe view SMS - admin console - lite version � Mozilla Firefox","")
	MouseClick("left",55,272,1)
	Sleep(2200)
	MouseClick("left",307,321,1)
	Sleep(250)
	Send("0424")
MouseClick("left",401,323,1)
Sleep(250)
Send($sCellValue)
Send("A56_23698")
MouseClick("left",973,317,1)
Sleep(250)
Send($sCellValue)
Sleep(250)
MouseClick("left",1214,322,1)
Sleep(250)
Send($sCellValue)
Sleep(250)
MouseClick("left",1375,322,1)
Sleep(250)
Send("0.00")
Sleep(250)
MouseClick("left",1832,473,1)
Sleep(2200)
	$i = $i +1
	WEnd
#endregion --- Au3Recorder generated code End ---
;_ExcelWriteSheetFromArray($oExcel2, $arr,1,1,0,0)
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
