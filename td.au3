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

_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\nbmctd.xls', 0, 0)
$i = 75
While $i < 88
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",863,547,1)
Sleep(1800)
MouseClick("left",755,245,1)
Sleep(900)
MouseClick("left",549,249,1)
MouseClick("left",549,316,1)
MouseClick("left",720,252,1)
Sleep(300)
Send($sCellValue)
MouseClick("left",958,243,1)
Sleep(300)
MouseClick("left",769,311,1)
Sleep(200)
MouseClick("left",853,584,1)
Sleep(800)
MouseClick("left",696,457,1)
Sleep(200)
MouseClick("left",505,201,1)
MouseClick("left",539,252,1)
MouseClick("left",548,271,1)
MouseClick("left",583,209,1)
Sleep(200)
MouseClick("left",354,269,1)
Sleep(2500)
MouseClick("left",646,206,1)
MouseClick("left",347,271,1)
Sleep(300)
MouseClick("left",331,158,1)
Sleep(7000)
MouseClick("left",684,464,1)
Sleep(200)
MouseClick("left",1075,140,1)
Sleep(2000)
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
