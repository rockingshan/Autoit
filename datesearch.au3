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

Local $oExcel = _ExcelBookOpen(@ScriptDir & '\564.xls', 0, 0)
$i = 481
While $i < 4485
	_WinWaitActivate("Modify Customer � Mozilla Firefox","")
	MouseClick("left",56,168,2)
	Sleep(150)
	Send("{CTRLDOWN}c{CTRLUP}")
	Sleep(150)
	$chkString = ClipGet()
	If $chkString == "Modify " Then
	ClipPut("")
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$arr[$i][0]=$sCellValue
	_WinWaitActivate("*new 1 - Notepad++","")
	Send($arr[$i][0])
	Send("{TAB}")
	_WinWaitActivate("Modify Customer � Mozilla Firefox","")
	Sleep(150)
	MouseClick("left",264,240,2)
	Sleep(200)
	Send($sCellValue)
	Sleep(250)
	Send("{ENTER}")
	Sleep(2500)
	MouseClick("left",265,938,1)
	Sleep(250)
	Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
	$arr[$i][1]=ClipGet()
	_WinWaitActivate("*new 1 - Notepad++","")
	Send($arr[$i][1])
	Send("{TAB}")
	_WinWaitActivate("Modify Customer � Mozilla Firefox","")
	ClipPut("")
	Sleep(200)
	MouseClick("left",279,973,1)
	Sleep(200)
	Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
	$arr[$i][2]=ClipGet()
	_WinWaitActivate("*new 1 - Notepad++","")
	Send($arr[$i][2])
	Send("{TAB}")
	_WinWaitActivate("Modify Customer � Mozilla Firefox","")
	ClipPut("")
	Sleep(150)
	MouseClick("left",278,1010,1)
	Sleep(150)
	Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
	$arr[$i][3]=ClipGet()
	_WinWaitActivate("*new 1 - Notepad++","")
	Send($arr[$i][3])
	Send("{TAB}")
	Send("{ENTER}")
	;_WinWaitActivate("Modify Customer � Mozilla Firefox","")
	ClipPut("")
	Sleep(150)
	
	$i = $i +1
Else
	MouseClick("left",1840,197,1)
	EndIf
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
	