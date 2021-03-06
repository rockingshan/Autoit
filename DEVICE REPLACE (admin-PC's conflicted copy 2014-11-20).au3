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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\ASSTAR.xls', 0, 0)
$i = 10
While $i < 20
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$wCellValue = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",433,168,1)
Sleep(100)
MouseClick("left",565,260,1)
MouseClick("left",565,308,1)
Sleep(150)
MouseClick("left",705,288,1)
Sleep(290)
MouseClick("left",613,241,1)
Sleep(120)
MouseClick("left",575,270,1)
MouseClick("left",762,251,1)
Sleep(100)
Send($sCellValue)
MouseClick("left",915,246,1)
Sleep(300)
MouseClick("left",500,305,1)
Sleep(100)
MouseClick("left",855,578,1)
Sleep(200)
MouseClick("left",853,555,1)
Sleep(110)
MouseMove(600,300)
MouseDown("left")
Sleep(180)
MouseMove(438,303)
MouseUp("left")
Sleep(110)
Send($wCellValue)
Sleep(170)
MouseClick("left",800,508,1)
Sleep(190)
MouseClick("left",325,163,1)
Sleep(15000)
MouseClick("left",650,458,1)
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

