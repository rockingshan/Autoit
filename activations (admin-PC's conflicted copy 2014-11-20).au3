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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\nbmc.xls', 0, 0)
$i = 29
While $i < 42
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue_1 = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",348,252,1)
MouseClick("left",568,356,1)
MouseClick("left",550,377,1)
Sleep(400)
MouseClick("left",423,257,1)
MouseClick("left",459,303,1)
Sleep(500)
MouseClick("left",436,464,1)
Sleep(900)
MouseClick("left",356,355,1)
MouseClick("left",350,455,1)
Sleep(400)
MouseClick("left",554,257,1)
MouseClick("left",548,303,1)
Sleep(400)
MouseClick("left",484,329,1)
MouseClick("left",406,349,2)
Sleep(400)
Send($sCellValue)
Sleep(400)
MouseClick("left",386,373,2)
Sleep(400)
Send($sCellValue_1)
Sleep(400)
MouseClick("left",582,350,2)
Send("1")
MouseClick("left",598,374,2)
Send("1")
MouseClick("left",643,260,1)
MouseClick("left",647,305,1)
Sleep(1200)
MouseMove(972,269)
MouseDown("left")
Sleep("400")
MouseMove(974,333)
MouseUp("left")
Sleep("400")
MouseClick("left",652,364,1)
Sleep("500")
MouseClick("left",857,615,1)
Sleep("300")
MouseClick("left",382,208,1)
Sleep(20000)
MouseClick("left",668,476,1)
MouseClick("left",423,201,1)


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
