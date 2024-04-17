#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#include <MsgBoxConstants.au3>
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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\channel.xlsx', 0, 0)
$i = 100
While $i < 140
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$sCellValue1 = _ExcelReadCell($oExcel, $i, 2)
	$sCellValue2 = _ExcelReadCell($oExcel, $i, 3)
	_WinWaitActivate("h8SSRMS - Google Chrome","")
	MouseClick("left",237,199,1)
	Sleep(250)
	MouseClick("left",247,224,1)
	Sleep(1500)
	MouseMove(169,347)
	MouseDown("left")
	MouseMove(10,354)
	MouseUp("left")
	Sleep(250)
	Send($sCellValue)
	Sleep(250)
	MouseClick("left",58,376,1)
	Sleep(4000)
	MouseClick("left",1783,319,1)
	Sleep(4000)
;~ 	MouseClick("left",1472,405,2)
;~ 	Sleep(250)
;~ 	Send("{DEL}")
;~ 	Sleep(250)
;~ 	Send($sCellValue1)
;~ 	MouseClick("left",1558,413,1)
;~ 	Sleep(1500)
	;zMsgBox($MB_SYSTEMMODAL, "Title", $sCellValue2, 10)
	MouseMove(628,357)
	MouseDown("left")
	MouseMove(476,351)
	MouseUp("left")
	Send("{DEL}")
	Sleep(250)
	Send($sCellValue1)
	Sleep(250)
	MouseMove(599,406)
	MouseDown("left")
	MouseMove(465,404)
	MouseUp("left")
	Send("{DEL}")
	Sleep(250)
	Send($sCellValue2)
	Sleep(250)
	MouseClick("left",789,419,1)
	Sleep(450)
	MouseClick("left",1730,236,1)
	Sleep(4000)
	$i = $i + 1
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
