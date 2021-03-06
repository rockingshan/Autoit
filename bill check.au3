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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\billtest.xls', 0, 0)
$i = 20
While $i < 40
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
_WinWaitActivate("Meghbela Digital SMS - Google Chrome","")
MouseClick("left",470,202,1)
Sleep(400)
Send($sCellValue)
Sleep(300)
Send("{ENTER}")
Sleep(3200)
MouseMove(683,569)
MouseDown("left")
MouseMove(861,566)
MouseUp("left")
MouseClick("left",998,310,1)
Sleep(400)
MouseClick("left",1120,228,1)
Sleep(6000)
_WinWaitActivate("Program Manager","")
MouseClick("left",732,785,1)
Msgbox(0,"Operation halted","Continue?")
_WinWaitActivate("Program Manager","")
Sleep(400)
MouseClick("left",360,785,1)
Sleep(300)
Send ( "!{F4}") 
#endregion --- Au3Recorder generated code End ---
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

