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

_WinWaitActivate("223.29.207.132 - Remote Desktop Connection","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\SRENTERPRISE.xls', 0, 0)
$i = 25
While $i < 50
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue_1 = _ExcelReadCell($oExcel, $j, 2)
MouseMove(431,155)
Sleep(350)
MouseDown("left")
Sleep(350)
MouseClick("left",431,155,1)
Send("^{A}")
Sleep(350)
Send($sCellValue_1)
Sleep(500)
MouseClick("left",434,196,1)
Sleep(400)
MouseMove(432,274)
Sleep(350)
MouseDown("left")
Sleep(350)
MouseClick("left",432,274,1)
Send("^{A}")
Sleep(350)
Send($sCellValue)
Sleep(500)
MouseClick("left",394,310,1)
Sleep(350)
MouseClick("left",302,474,1)
Sleep(350)
MouseClick("left",372,421,1)
Sleep(350)
;~ _WinWaitActivate("Untitled - Notepad","")
;~ Send($sCellValue_1)
;~ Send($sCellValue)
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
