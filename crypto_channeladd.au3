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

_WinWaitActivate("Cryptoguard | Meghbela Digital - Google Chrome","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\cryptochanneladd.xls', 0, 0)
$i = 220
While $i < 379
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseClick("left",1208,489,1)
Sleep(300)
MouseClick("left",1186,535,1)
Sleep(1800)
MouseClick("left",759,229,1)
Sleep(150)
MouseClick("left",487,268,1)
Sleep(150)
MouseClick("left",483,256,1)
Sleep(300)
Send($sCellValue)
Sleep(300)
MouseMove(1196,294)
MouseDown("left")
MouseMove(1159,590)
MouseUp("left")
Sleep(300)
MouseClick("left",492,646,1)
Sleep(2100)
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
