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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\pack.xlsx', 0, 0)
$i = 4
While $i < 13
	_WinWaitActivate("Price Plan - Google Chrome","")
	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	$sCellValue1 = _ExcelReadCell($oExcel, $i, 2)
;~ 	$sCellValue2 = _ExcelReadCell($oExcel, $i, 2)
;~ 	$sCellValue3 = _ExcelReadCell($oExcel, $i, 4)
	MouseClick("left",483,263,2)
	Sleep(250)
	Send($sCellValue)
	Sleep(500)
	MouseClick("left",623,271,1)
	Sleep(2000)
	Send($sCellValue1)
	Sleep(800)
	MouseClick("left",504,303,1)
	Sleep(250)
	Send($sCellValue1)
	Sleep(250)
	Send("{ENTER}")
	Sleep(250)
	MouseClick("left",1198,458,1)
	Sleep(250)
	MouseClick("left",443,415,1)
	MouseClick("left",405,458,1)
	Sleep(250)
	MouseClick("left",665,696,1)
	Sleep(250)
	MouseClick("left",262,780,1)
	Sleep(650)
	MouseClick("left",1066,167,1)
	Sleep(250)
	MouseClick("left",665,776,1)
	Sleep(250)
	MouseClick("left",665,807,1)
	Sleep(250)
	MouseClick("left",1464,807,1)
	Sleep(250)
	MouseClick("left",1462,833,1)
	Sleep(250)
	MouseMove(1837,808)
	MouseDown("left")
	MouseMove(1839,900)
	MouseUp("left")
	Sleep(250)
	MouseClick("left",1459,907,1)
	Sleep(250)
	MouseMove(1037,628)
	MouseDown("left")
	MouseMove(1034,708)
	MouseUp("left")
	Sleep(250)
	MouseClick("left",1497,979,1)
	Sleep(5000)
;~ 	MouseClick("left",46,668,1)
;~ 	Sleep(1500)
	;_WinWaitActivate("Define Pricing Details - Google Chrome","")
;~ 	Sleep(650)
;~ 	MouseClick("left",1078,285,1)
;~ 	Sleep(250)
;~ 	MouseClick("left",1167,345,1)
;~ 	Sleep(250)
;~ 	MouseClick("left",1167,525,1)
;~ 	Sleep(250)
;~ 	MouseClick("left",1180,567,1)
;~ 	Sleep(250)
;~ 	MouseClick("left",362,370,1)
;~ 	Sleep(2500)
;~ 	MouseMove(852,750)
;~ 	Sleep(250)
;~ 	MouseDown("left")
;~ 	MouseMove(1225,758)
;~ 	MouseUp("left")
;~ 	MouseClick("left",1173,509,1)
;~ 	Send($sCellValue3/2)
;~ 	Sleep(250)
;~ 	MouseClick("left",1131,568,1)
;~ 	Send($sCellValue3*6)
;~ 	Sleep(250)
;~ 	MouseClick("left",1128,609,1)
;~ 	Send($sCellValue3)
;~ 	Sleep(250)
;~ 	MouseClick("left",1136,654,1)
;~ 	Send($sCellValue3*2)
;~ 	Sleep(250)
;~ 	MouseClick("left",1136,694,1)
;~ 	Send($sCellValue3*0.23)
;~ 	Sleep(250)
;~ 	MouseMove(1314,552)
;~ 	MouseDown("left")
;~ 	MouseMove(1310,717)
;~ 	MouseUp("left")
;~ 	MouseClick("left",1130,722,1)
;~ 	Send($sCellValue3*3)
;~ 	Sleep(250)
	MouseClick("left",1728,820,1)
	Sleep(3000)
	MouseClick("left",71,635,1)
	Sleep(2000)
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
