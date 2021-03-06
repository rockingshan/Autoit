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

_WinWaitActivate("DVTCASAPP - VNC Viewer","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\Temp1.xls', 0, 0)
$i = 35
While $i < 50
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseMove(416,159)
MouseDown("left")
MouseMove(263,157)
MouseUp("left")
Send($sCellValue)
MouseClick("left",873,167,1)
Sleep(750)
MouseClick("left",684,266,1)
MouseClick("left",841,510,1)
MouseClick("left",473,412,1)
Sleep(950)
_WinWaitActivate("Untitled - Notepad","")
	Send($sCellValue & ",")
	_WinWaitActivate("DVTCASAPP - VNC Viewer","")
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
