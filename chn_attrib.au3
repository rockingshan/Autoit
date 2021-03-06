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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\chn_attrib.xls', 0, 0)
$i = 2
While $i < 28
   $j = $i
   $sCellValue = _ExcelReadCell($oExcel, $j, 1)
   MouseClick("left",610,298,1)
Sleep(250)
Send($sCellValue)
Sleep(150)
Send("{TAB}")
Sleep(250)
MouseClick("left",424,222,1)
MouseClick("left",716,303,1)
Sleep(200)
MouseClick("left",740,351,1)
Sleep(500)
MouseClick("left",447,222,1)
Sleep(200)
MouseClick("left",328,188,1)
Sleep(190)
MouseClick("left",678,465,1)
Sleep(150)

	_WinWaitActivate("Meghbela Digital SMS - Google Chrome","")
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

