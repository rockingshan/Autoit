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

_WinWaitActivate("192.168.1.150 - Remote Desktop Connection","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\CAL AMN.xlsx', 0, 0)
$i = 50
While $i < 127
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseMove(546,171)
MouseDown("left")
MouseMove(369,162)
MouseUp("left")
Sleep(200)
Send($sCellValue)
Sleep(200)
MouseClick("left",1281,172,1)
Sleep(950)
MouseClick("left",497,264,1)
Send("{CTRLDOWN}")
MouseClick("left",479,298,1)
Send("{CTRLUP}")
MouseClick("left",1280,544,1)
Sleep(300)
MouseClick("left",669,447,1)
Sleep(1050)
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
