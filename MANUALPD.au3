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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\NEXUS.xls', 0, 0)
$i = 7
While $i < 20
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseClick("left",441,175,1)
Sleep(175)
MouseClick("left",382,178,1)
Sleep(350)
MouseClick("left",498,251,1)
Sleep(200)
MouseClick("left",523,319,1)
MouseClick("left",719,255,1)
Send($sCellValue)
Sleep(150)
MouseClick("left",940,257,1)
Sleep(550)
MouseClick("left",748,314,1)
MouseClick("left",838,588,1)
Sleep(250)
MouseClick("left",491,219,1)
Sleep(150)
MouseClick("left",524,260,1)
Sleep(80)
MouseClick("left",543,333,1)
MouseClick("left",577,216,1)
MouseClick("left",353,278,1)
Sleep(7100)
MouseClick("left",648,219,1)
Sleep(200)
MouseClick("left",346,280,1)
MouseClick("left",314,173,1)
Sleep(8000)
Send("{ENTER}")
Sleep(200)
MouseClick("left",373,216,1)
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

