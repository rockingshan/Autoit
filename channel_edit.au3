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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\extend.xlsx', 0, 0)
$i = 5
While $i < 36
_WinWaitActivate("Validity Extension - Google Chrome","")
Sleep(300)
MouseClick("left",314,312,1)
Sleep(300)
$sCellValue = _ExcelReadCell($oExcel, $i, 1)
Send($sCellValue)
Sleep(300)
Send("{ENTER}")
Sleep(3000)
MouseClick("left",331,349,1)
Sleep(300)
Send("15{ENTER}")
Sleep(300)
MouseClick("left",1194,354,1)
Sleep(300)
Send("da{ENTER}")
Sleep(300)
MouseClick("left",1692,406,1)
Sleep(4000)
MouseClick("left",1828,404,1)
Sleep(4000)
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
