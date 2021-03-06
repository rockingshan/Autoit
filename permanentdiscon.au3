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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\PD.xlsx', 0, 0)
$i = 3
While $i < 30
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseClick("left",748,255,1)
Sleep(175)
MouseClick("left",549,247,1)
Sleep(150)
MouseClick("left",516,316,1)
Sleep(175)
MouseClick("left",716,254,1)
Sleep(150)
Send($sCellValue)
Sleep(150)
MouseClick("left",937,246,1)
Sleep(3000)
MouseClick("left",688,310,1)
MouseClick("left",849,580,1)
Sleep(1100)
MouseClick("left",515,213,1)
Sleep(150)
MouseClick("left",527,256,1)
Sleep(75)
MouseClick("left",545,326,1)
MouseClick("left",580,207,1)
MouseClick("left",343,272,1)
Sleep(5500)
MouseClick("left",640,213,1)
MouseClick("left",350,271,1)
Sleep(550)
MouseClick("left",311,178,1)
Sleep(4000)
Send("{ENTER}")
Sleep(250)
MouseClick("left",410,179,1)
Sleep(150)
MouseClick("left",366,216,1)
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

