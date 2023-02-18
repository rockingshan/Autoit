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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\pack.xlsx', 0, 0)
$i = 50
While $i < 70
	_WinWaitActivate("Bundle - Google Chrome","")
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$sCellValue1 = _ExcelReadCell($oExcel, $i, 2)
	MouseClick("left",491,249,2)
	Sleep(250)
	Send($sCellValue)
	Sleep(500)
	MouseClick("left",733,249,1)
	Sleep(2000)
	MouseClick("left",516,288,1)
	Sleep(200)
	Send($sCellValue1)
	Sleep(200)
	MouseClick("left",649,296,1)
	Sleep(250)
	MouseClick("left",524,330,1)
	Sleep(250)
	MouseClick("left",526,361,1)
	Sleep(800)
	MouseClick("left",1767,409,1)
	Sleep(1800)
	Send("{TAB}")
	Sleep(250)
	Send($sCellValue1)
	Sleep(250)
	MouseClick("left",1848,424,1)
	Sleep(850)
	MouseClick("left",1741,746,1)
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
