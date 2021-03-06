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
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\silver.xlsx', 0, 0)
$i = 16
While $i < 36
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)

MouseClick("left",510,255,1)
Sleep(150)
Send($sCellValue)
Sleep(150)
Send("{ENTER}")
Sleep(6000)
MouseClick("left",920,309,1)
Sleep(100)
MouseClick("left",870,594,1)
Sleep(100)
MouseClick("left",460,379,3)
Send("si")
MouseClick("left",448,405,1)
Sleep(100)
MouseMove(495,606)
MouseDown("left")
Sleep(150)
MouseMove(760,599)
MouseUp("left")
Sleep(100)
MouseClick("left",676,379,2)
MouseClick("left",626,382,1)
MouseClick("left",673,400,1)
Sleep(200)
MouseMove(732,604)
Sleep(250)
MouseDown("left")
MouseMove(425,607)
Sleep(250)
MouseUp("left")
MouseClick("left",313,176,1)
Sleep(15000)
MouseClick("left",685,466,1)
Sleep(250)
MouseClick("left",409,183,1)


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
