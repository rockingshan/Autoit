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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\subs.xlsx', 0, 0)
$i = 500
While $i < 654
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	_WinWaitActivate("Securetv DRM - Google Chrome","")
	MouseClick("left",470,125,1)
	Sleep(800)
	MouseClick("left",472,271,1)
	Sleep(500)
	Send($sCellValue)
	Sleep(800)
	MouseClick("left",1858,390,1)
	Sleep(500)
	MouseMove(1016,750)
	MouseDown("left")
	Sleep(250)
	MouseMove(1015,750)
	MouseUp("left")
	Sleep(250)
	MouseMove(773,914)
	MouseDown("left")
	Sleep(250)
	MouseMove(772,914)
	MouseUp("left")
	Sleep(500)
	MouseClick("left",1848,958,1)
	Sleep(1000)
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
