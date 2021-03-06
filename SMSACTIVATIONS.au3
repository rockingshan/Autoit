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

_WinWaitActivate("DuoSubscribe Version 5.0 - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\MUNNA1.xls', 0, 0)
$i = 23
While $i < 30
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue_1 = _ExcelReadCell($oExcel, $j, 2)
	$sCellValue_2 = _ExcelReadCell($oExcel, $j, 3)
MouseMove(810,298)
Sleep(150)
MouseDown("left")
MouseMove(388,298)
Sleep(150)
MouseUp("left")
Send($sCellValue)
Send(" ")
Send($sCellValue_1)
Send(" STKTSUMA ")
Send($sCellValue_2)
Send(" ")
Sleep(200)
MouseClick("left",852,299,1)
Sleep(13000)
Send("{ENTER}")
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
