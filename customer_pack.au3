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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\hi8pack.xlsx', 0, 0)
$i = 1
While $i < 30
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$sCellValue1 = _ExcelReadCell($oExcel, $i, 2)
	_WinWaitActivate("h8SSRMS - Google Chrome","")
	MouseClick("left",104,190,1)
	Sleep(500)
	MouseClick("left",88,265,1)
	Sleep(2500)
	MouseMove(137,361)
	MouseDown("left")
	MouseMove(20,362)
	MouseUp("left")
	Send($sCellValue)
	Sleep(800)
	MouseClick("left",60,411,1)
	Sleep(3500)
	MouseClick("left",436,339,1)
	Sleep(3500)
	;wait for box
;~ 	MouseClick("left",1472,405,2)
;~ 	Sleep(250)
;~ 	Send("{DEL}")
;~ 	Sleep(250)
;~ 	Send($sCellValue1)
;~ 	MouseClick("left",1558,413,1)
;~ 	Sleep(1500)
	MsgBox($MB_OK, "Title", "Are you done")
	_WinWaitActivate("h8SSRMS - Google Chrome","")
	MouseClick("left",67,504,1)
	Sleep(1000)
	MouseClick("left",67,504,1)
	Sleep(2000)
	MouseClick("left",725,435,1)
	Sleep(250)
	MouseMove(725,464)
	MouseDown("left")
	MouseMove(730,470)
	MouseUp("left")
	Sleep(2500)
	MouseClick("left",731,470,1)
	Sleep(500)
	Send($sCellValue1)
	Sleep(250)
	;MouseClick("left",723,539,1)
	Send("{ENTER}")
	Sleep(2500)
	MouseClick("left",929,633,1)
	Sleep(3500)
	MouseClick("left",986,634,1)
	Sleep(3500)
	MouseClick("left",986,634,1)
	Sleep(3500)
	MouseClick("left",980,740,1)
	Sleep(3500)
	MouseClick("left",969,758,1)
	Sleep(8000)
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
