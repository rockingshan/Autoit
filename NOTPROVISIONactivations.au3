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

_WinWaitActivate("DuoSubscribe Version 5.0 - Google Chrome","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\JAIMATADI.xlsx', 0, 0)
$i = 6
While $i < 42
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",622,307,1)
MouseClick("left",560,334,1)
Sleep(400)
MouseClick("left",441,208,1)
MouseClick("left",454,238,1)
Sleep(400)
Send("pl")
MouseClick("left",446,265,1)
Sleep(500)
MouseClick("left",352,298,1)
MouseClick("left",348,322,1)
Sleep(400)
MouseClick("left",525,210,1)
MouseClick("left",526,235,1)
Send($sCellValue)
MouseClick("left",631,240,1)
Sleep(400)
MouseClick("left",348,294,1)
MouseClick("left",693,209,1)
Sleep(400)
MouseClick("left",310,298,1)
MouseClick("left",570,258,1)
Sleep(400)
Send("JAIMATADI_CABLE")
Send("{ENTER}")
Sleep(500)
MouseClick("left",388,165,1)
Sleep("8000")
MouseClick("left",682,460,1)
Sleep(400)
MouseClick("left",443,164,1)
MouseClick("left",383,207,1)


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
