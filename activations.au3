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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\MULTIMEDIA199ACTIV.xls', 0, 0)
$i = 85
While $i < 201
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue_1 = _ExcelReadCell($oExcel, $j, 2)
MouseClick("left",437,175,1)
MouseClick("left",473,313,1)
MouseMove(475,339)
MouseDown("left")
MouseMove(476,339)
MouseUp("left")
Sleep(400)
MouseClick("left",430,218,1)
MouseMove(441,252)
MouseDown("left")
Sleep(500)
MouseClick("left",441,252,1)
Send("pr")
MouseClick("left",442,342,1)
Sleep(900)
MouseClick("left",352,304,1)
MouseClick("left",352,473,1)
Sleep(400)
MouseClick("left",596,214,1)
MouseClick("left",546,251,1)
Sleep(400)
MouseClick("left",509,275,1)
MouseClick("left",342,303,2)
Sleep(400)
Send($sCellValue)
Sleep(400)
MouseClick("left",372,327,2)
Sleep(400)
Send($sCellValue_1)
Sleep(400)
MouseClick("left",570,305,2)
Send("1")
MouseClick("left",567,331,2)
Send("1")
MouseClick("left",671,214,1)
MouseClick("left",658,268,1)
Sleep("800")
MouseMove(986,261)
MouseDown("left")
Sleep("400")
MouseMove(1011,524)
MouseUp("left")
Sleep("400")
MouseClick("left",734,531,1)
Sleep("500")
MouseClick("left",863,602,1)
Sleep("300")
MouseClick("left",360,168,1)
Sleep("8000")
MouseClick("left",669,470,1)
MouseClick("left",385,217,1)


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
