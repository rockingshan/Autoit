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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\customer.xlsx', 0, 0)
$i = 14
While $i < 28
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	;_WinWaitActivate("h8SSRMS - Google Chrome","")
	MouseClick("left",85,187,1)
    MouseClick("left",86,259,1)
	Sleep(5000)
	MouseClick("left",142,360,2)
	Sleep(250)
	Send($sCellValue)
	Sleep(250)
	MouseClick("left",34,403,1)
	Sleep(2000)
	MouseClick("left",442,341,1)
	Sleep(9000)
	;zMsgBox($MB_SYSTEMMODAL, "Title", $sCellValue2, 10)
	MouseClick("left",66,507,1)
	Sleep(1200)
   MouseClick("left",66,507,1)
	Sleep(2000)
	MouseClick("left",750,437,1)
	Sleep(800)
	MouseClick("left",753,482,1)
	Sleep(3000)
	MouseClick("left",755,465,1)
	Send("100  {ENTER}")
	Sleep(8000)
	MouseClick("left",942,632,1)
	Sleep(3000)
    MouseClick("left",983,631,1)
    Sleep(3000)
    MouseClick("left",983,631,1)
    Sleep(3000)
    MouseClick("left",974,736,1)
    Sleep(3000)
    MouseClick("left",974,736,1)
	Sleep(15000)
	MouseClick("left",75,450,1)
	Sleep(250)
    MouseClick("left",79,429,1)
	Sleep(250)
	MouseClick("left",722,540,1)
	Send("inac{ENTER}")
	Sleep(250)
	MouseClick("left",742,583,1)
	Send("ss")
	Sleep(500)
	MouseClick("left",937,612,1)
	Sleep(5000)
	Send("{PGDN}")
	Sleep(250)
	MouseClick("left",1798,967,1)
	Sleep(2000)
	MouseClick("left",928,655,1)
	Sleep(9000)
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
