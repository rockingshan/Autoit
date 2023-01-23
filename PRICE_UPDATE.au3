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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\564.xls', 0, 0)
$i = 1
While $i < 9
	_WinWaitActivate("Operator Plan wise Price — Mozilla Firefox","")
	MouseClick("left",1260,210,2)
	Sleep(150)
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	Send($sCellValue)
	Sleep(250)
	Send("{ENTER}")
	Sleep(2800)
	;comment from here
;~ 	MouseClick("left",365,239,1)
;~ 	Sleep(250)
;~ 	MouseClick("left",350,285,1)
;~ 	Sleep(250)
;~ 	MouseClick("left",1165,239,1)
;~ 	Send("bronze{ENTER}")
	;comment end here
	MouseClick("left",1754,287,1)
	Sleep(3050)
	MouseClick("left",1770,407,1)
	Sleep(3800)
	MouseClick("left",1249,606,1)
	Sleep(800)
	MouseClick("left",1178,668,1)
	Sleep(800)
	MouseClick("left",1748,647,1)
	Sleep(3500)
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
	