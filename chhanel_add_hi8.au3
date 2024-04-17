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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\channel_create.xlsx', 0, 0)
$i = 140
$j = 749
While $i < 173
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$sCellValue1 = _ExcelReadCell($oExcel, $i, 2)
	$sCellValue2 = _ExcelReadCell($oExcel, $i, 3)
	$sCellValue3 = _ExcelReadCell($oExcel, $i, 4)
	;_WinWaitActivate("h8SSRMS - Google Chrome","")
	MouseClick("left",44,263,1)
	Sleep(3000)
	MouseClick("left",534,318,1)
	Sleep(250)
	Send($sCellValue)
	Sleep(250)
	MouseClick("left",913,363,1)
	Sleep(250)
	MouseClick("left",1386,307,1)
	Sleep(500)
	Send($sCellValue2)
	;zMsgBox($MB_SYSTEMMODAL, "Title", $sCellValue2, 10)
	Send("{ENTER}")
	Sleep(250)
	MouseClick("left",1408,372,1)
	Send($sCellValue1)
	Send("{ENTER}")
	MouseClick("left",1281,392,1)
	Sleep(3000)
	MouseClick("left",625,334,1)
	Send($j)
	Sleep(250)
	MouseClick("left",1404,390,2)
	Send($sCellValue3)
	Sleep(250)
	MouseClick("left",1402,417,1)
	Sleep(3000)
	MouseClick("left",1407,449,1)
	MouseClick("left",1411,492,1)
	Sleep(1500)
	MouseClick("left",576,401,1)
	Send("C")
	Send($j)
	Sleep(250)
	MouseClick("left",573,446,1)
	Sleep(250)
	MouseClick("left",564,482,1)
	Sleep(250)
	MouseClick("left",1739,226,1)
	Sleep(5000)
	$i = $i +1
	$j = $j +1
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
