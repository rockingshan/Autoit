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

_WinWaitActivate("Safe View - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\refresh.xlsx', 0, 0)
$i = 5
While $i < 6
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 4)
;~ 	$sCellValue_1 = _ExcelReadCell($oExcel, $j, 5)
MouseClick("left",282,176,1)
Sleep(300)
MouseClick("left",308,274,1)
Sleep(300)
Send($sCellValue)
MouseClick("left",457,271,1)
Sleep(300)
MouseClick("left",234,267,1)
Sleep(200)
MouseClick("left",544,380,2)
Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
Sleep(600)
Send("9006178")
MouseClick("left",660,391,1)
Sleep(700)
MouseClick("left",557,387,2)
Send("{CTRLDOWN}v{CTRLUP}")
MouseClick("left",652,384,1)
Sleep(800)
MouseClick("left",256,374,1)
Sleep(100)
MouseClick("left",454,279,1)
Sleep(5400)
ClipPut("")
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
