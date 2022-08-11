#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#region --- Internal functions Au3Recorder Start ---
Global $arr[9999][4]
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
;;;80% zoom for modify customer
_AU3RecordSetup()
#endregion --- Internal functions Au3Recorder End ---
;Local $oExcel2 = _ExcelBookNew()
Run ("C:\Program Files (x86)\Notepad++\notepad++.exe")

Local $oExcel = _ExcelBookOpen(@ScriptDir & '\mdch047.xls', 0, 0)
$i = 1
While $i < 481
	_WinWaitActivate("History Details — Mozilla Firefox","")
	Sleep(150)
	MouseClick("left",1360,340,1)
	Sleep(150)
	MouseClick("left",607,235,2)
	Sleep(150)
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	Send($sCellValue)
	Sleep(250)
	Send("{ENTER}")
	Sleep(10500)
;~ 	MouseClick("left",265,938,1)
;~ 	Sleep(250)
MouseMove(261,334)
MouseDown("left")
MouseMove(99,334)
MouseUp("left")
Sleep(150)
	Send("10/07/2022")
	Sleep(150)
	MouseClick("left",260,390,2)
	Send("mb102{ENTER}")
	Sleep(8250)
	MouseClick("left",1149,365,1)
	Sleep(250)
	Send("modify{SPACE}cu{ENTER}")
	Sleep(550)
	MouseClick("left",1745,432,1)
	Sleep(15250)
	MouseClick("left",1388,546,1)
	Sleep(20250)
	MouseClick("left",1100,603,1)
	Sleep(650)
	_WinWaitActivate("*new 1 - Notepad++","")
	Send($sCellValue)
	_WinWaitActivate("History Details — Mozilla Firefox","")
	MouseClick("left",1100,603,1)
	Sleep(650)
	Send("{CTRLDOWN}ac{CTRLUP}")
	_WinWaitActivate("*new 1 - Notepad++","")
	Send("{CTRLDOWN}v{CTRLUP}")
	Send("{ENTER}")
		;_WinWaitActivate("Modify Customer — Mozilla Firefox","")
	ClipPut("")
	Sleep(150)
	
	$i = $i +1
   WEnd
#endregion --- Au3Recorder generated code End ---
;_ExcelWriteSheetFromArray($oExcel2, $arr,1,1,0,0)
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
	