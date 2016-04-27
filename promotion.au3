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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\CODEADD.xlsx', 0, 0)
$i = 2
While $i < 18
   $j = $i
   $sCellValue = _ExcelReadCell($oExcel, $j, 1)

MouseMove(1043,479)
Sleep(500)
MouseDown("left")
Sleep(500)
MouseMove(1039,583)
Sleep(500)
MouseUp("left")
Sleep(500)
MouseClick("left",415,564,1)
Send("{ENTER}")
Sleep(800)
MouseMove(1040,474)
Sleep(500)
MouseDown("left")
Sleep(500)
MouseMove(1040,573)
MouseUp("left")
Sleep(500)
MouseClick("left",448,555,2)
Sleep(500)
Send($sCellValue)
Sleep(500)
Send("{TAB}")
Sleep(500)

;~    MouseClick("left",436,560,1)
;~ Send("{ENTER}")
;~ Sleep(450)
;~ MouseMove(1005,478)
;~ MouseDown("left")
;~ MouseMove(994,578)
;~ MouseUp("left")
;~ Sleep(500)
;~ MouseMove(611,584)
;~ MouseDown("left")
;~ MouseMove(901,593)
;~ MouseUp("left")
;~ Sleep(500)
;~ MouseClick("left",848,562,1)
;~ Sleep(300)
;~ MouseMove(1006,478)
;~ MouseDown("left")
;~ MouseMove(1000,571)
;~ MouseUp("left")
;~ Sleep(300)
;~ MouseMove(924,579)
;~ MouseDown("left")
;~ MouseMove(879,584)
;~ MouseUp("left")
;~ Sleep(500)
;~ MouseMove(871,584)
;~ MouseDown("left")
;~ MouseMove(862,584)
;~ MouseUp("left")
;~ Sleep(400)
;~ MouseMove(842,584)
;~ MouseDown("left")
;~ MouseMove(543,576)
;~ MouseUp("left")
;~ Sleep(400)
;~ MouseClick("left",417,567,2)
;~ Sleep(300)
;~ Send($sCellValue)
;~ Sleep(150)
;~ Send("{TAB}")
;~ Sleep(500)
;~ MouseMove(1008,475)
;~ MouseDown("left")
;~ MouseMove(1004,568)
;~ MouseUp("left")

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

