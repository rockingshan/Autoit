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





#region --- Au3Recorder generated code Start (v3.3.9.5 KeyboardLayout=00000409)  ---
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\CODEADD.xlsx', 0, 0)
$i = 4
While $i < 18
   $j = $i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$sCellValue_1 = _ExcelReadCell($oExcel, $j, 3)
	$sCellValue_2 = _ExcelReadCell($oExcel, $j, 4)
	$sCellValue_3 = _ExcelReadCell($oExcel, $j, 5)
	$sCellValue_4 = _ExcelReadCell($oExcel, $j, 6)
MouseClick("left",596,284,1)
Sleep(350)
Send ($sCellValue)
Sleep(350)
Send("{BACKSPACE}{DOWN}{ENTER}")
MouseClick("left",582,212,1)
Sleep(350)
;~ Send("{ALTDOWN}{ALTUP}")
;~ _WinWaitActivate("E:\Dropbox\autoit scripts\CODEADD.au3 * SciTE [4 of 4]","")
;~ Send("{ALTDOWN}{ALTUP}")
;~ _WinWaitActivate("Program Manager","")
;~ MouseClick("left",532,747,1)
;~ _WinWaitActivate("classname=TaskListThumbnailWnd","")
;~ MouseClick("left",851,165,1)
;~ _WinWaitActivate("Microsoft Excel - Book1","")
;~ MouseClick("left",393,224,1)
;~ Send("{CTRLDOWN}c{CTRLUP}{ALTDOWN}{ALTUP}")
;~ _WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseClick("left",658,507,1)
Sleep(350)
Send($sCellValue_1)
Sleep(350)
;~ _WinWaitActivate("Microsoft Excel - Book1","")
;~ MouseClick("left",489,223,1)
;~ Send("{CTRLDOWN}c{CTRLUP}{ALTDOWN}{ALTUP}")
;~ _WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseClick("left",654,479,1)
Send($sCellValue_2)
Sleep(350)
;~ _WinWaitActivate("Microsoft Excel - Book1","")
;~ MouseClick("left",654,225,1)
;~ Send("{CTRLDOWN}c{CTRLUP}{ALTDOWN}{ALTUP}")
;~ _WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseClick("left",676,346,1)
Send($sCellValue_3)
Sleep(350)
;~ _WinWaitActivate("Microsoft Excel - Book1","")
;~ MouseClick("left",728,226,1)
;~ Send("{CTRLDOWN}c{CTRLUP}{ALTDOWN}{ALTUP}")
;~ _WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseClick("left",505,213,1)
Sleep(450)
MouseClick("left",580,485,1)
Send($sCellValue_4)
MouseClick("left",898,240,1)
Sleep(450)
MouseClick("left",433,237,1)
MouseClick("left",311,170,1)
Sleep(850)
MouseClick("left",648,469,1)
#endregion --- Au3Recorder generated code End ---
$i = $i +1
   WEnd


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




