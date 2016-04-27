#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#region --- Internal functions Au3Recorder Start ---
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
_WinWaitActivate("Meghbela Digital SMS - Google Chrome","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\packagedetail.xls', 0, 0)
$i = 10
While $i < 28
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue_4 = _ExcelReadCell($oExcel, $j, 4)
	$sCellValue_5 = _ExcelReadCell($oExcel, $j, 5)
	$sCellValue_6 = _ExcelReadCell($oExcel, $j, 6)
	MouseClick("left",610,298,1)
	Sleep(250)
	Send($sCellValue)
	Sleep(150)
	Send("{TAB}")
	Sleep(250)
	MouseMove(578,506)
	MouseDown("left")
	MouseMove(577,506)
	MouseUp("left")
	Send($sCellValue_6)
	Sleep(240)
	MouseMove(776,510)
	MouseDown("left")
	MouseMove(835,515)
	MouseUp("left")
	Sleep(250)
	Send("1")
	MouseClick("left",584,223,1)
	Sleep(300)
	MouseMove(791,477)
	MouseDown("left")
	MouseMove(639,470)
	MouseUp("left")
	Send($sCellValue_5)
	Sleep(250)
	MouseMove(749,501)
	MouseDown("left")
	MouseMove(631,494)
	MouseUp("left")
	Send($sCellValue_4)
	MouseClick("left",875,228,1)
	Sleep(250)
	MouseClick("left",563,351,1)
	MouseClick("left",574,326,1)
	Sleep(300)
	MouseClick("left",559,376,1)
	MouseMove(428,256)
	MouseDown("left")
	Sleep(350)
	MouseMove(428,257)
	MouseUp("left")
	MouseClick("left",328,188,1)
	Sleep(450)
	MouseClick("left",660,469,1)

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
#endregion --- Au3Recorder generated code End ---



