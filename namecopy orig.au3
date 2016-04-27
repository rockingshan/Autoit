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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\accountlistmod.xls', 0, 0)
$i = 2
While $i < 4
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseClick("left",476,212,1)
Sleep(75)
Send($sCellValue)
Sleep(150)
Send("{ENTER}")
Sleep(200)
MouseMove(724,333)
MouseDown("left")
MouseMove(279,306)
MouseUp("left")
Send("{CTRLDOWN}c{CTRLUP}")
Sleep(200)
_WinWaitActivate("Untitled - Notepad","")
Send($sCellValue)
Send("{TAB}")
Send("{CTRLDOWN}v{CTRLUP}")
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseMove(688,359)
MouseDown("left")
MouseMove(341,345)
MouseUp("left")
Send("{CTRLDOWN}c{CTRLUP}")
Sleep(200)
_WinWaitActivate("Untitled - Notepad","")
Send("{TAB}")
Send("{CTRLDOWN}v{CTRLUP}")
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseMove(443,576)
MouseDown("left")
MouseMove(1257,593)
MouseUp("left")
Send("{CTRLDOWN}c{CTRLUP}")
Sleep(200)
_WinWaitActivate("Untitled - Notepad","")
Send("{TAB}")
Send("{CTRLDOWN}v{CTRLUP}")
Send("{ENTER}")
Sleep(150)
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseMove(373,174)
MouseDown("left")
MouseMove(374,174)
MouseUp("left")
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

