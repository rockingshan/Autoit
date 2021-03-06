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
_WinWaitActivate("Meghbela - Google Chrome","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\cwtlst.xls', 0, 0)
$i = 199
While $i < 217
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue2 = _ExcelReadCell($oExcel, $j, 3)
	$sCellValue3 = _ExcelReadCell($oExcel, $j, 4)
MouseMove(753,152)
Sleep(200)
MouseMove(594,302)
Sleep(1000)
MouseClick("left",594,302,1)
Sleep(3050)
MouseClick("left",565,246,1)
Sleep(250)
Send("{CTRLDOWN}a{CTRLUP}")
Sleep(555)
;~ MsgBox(4096, "", $sCellValue)
Send($sCellValue)
Sleep(750)
MouseClick("left",497,369,1)
Sleep(450)
Send($sCellValue2)
Sleep(950)
MouseClick("left",979,369,3)
Sleep(500)
Send($sCellValue3)
Sleep(500)
MouseClick("left",1029,486,1)
Sleep(2500)
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
