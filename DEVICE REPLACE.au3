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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\REPLACE.xls', 0, 0)
$i = 19
While $i < 32
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$wCellValue = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",438,213,1)
Sleep(100)
MouseClick("left",542,253,1)
MouseClick("left",541,302,1)
Sleep(150)
MouseClick("left",714,286,1)
Sleep(290)
MouseClick("left",578,238,1)
Sleep(120)
MouseClick("left",574,271,1)
MouseClick("left",763,244,1)
Sleep(100)
Send($sCellValue)
MouseClick("left",911,239,1)
Sleep(300)
MouseClick("left",504,297,1)
Sleep(160)
MouseClick("left",838,569,1)
Sleep(200)
MouseClick("left",846,536,1)
Sleep(190)
MouseMove(593,304)
MouseDown("left")
Sleep(180)
MouseMove(413,292)
MouseUp("left")
Sleep(190)
Send($wCellValue)
Sleep(210)
MouseClick("left",828,499,1)
Sleep(210)
MouseClick("left",326,165,1)
Sleep(10000)
MouseClick("left",683,447,1)
Sleep(130)
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

