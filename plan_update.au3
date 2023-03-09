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
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\plan_old.xlsx', 0, 0)
$i = 40
;_WinWaitActivate("Price Plan - Google Chrome","")
;_WinWaitActivate("Define Pricing Details - Google Chrome","")
;_WinWaitActivate("prc_plan - Google Chrome","")
_WinWaitActivate("def_pricing_dtls - Google Chrome","")

While $i < 69

	$sCellValue = _ExcelReadCell($oExcel, $i, 1)
	;Enter plan code
;~ 	MouseClick("left",484,258,2)
;~ 	Sleep(250)
;~ 	Send($sCellValue)
;~ 	Sleep(500)
;~ 	Send("{ENTER}")
;~ 	Sleep(3000)
;~ 	MouseClick("left",444,377,1)
;~ 	Send("NTO{SPACE}2{SPACE}Package")
	;FOLLOWING BLOCK CHANGES CONFIG
	;Change active Inactive
;~ 	MouseClick("left",1609,258,1)
;~ 	Sleep(200)
;~ 	Send("In{ENTER}")
;~ 	Sleep(250)
;~ 	MouseClick("left",262,782,1)
;~ 	Sleep(1000)
;~ 	MouseClick("left",1077,169,1)
;~ 	Sleep(500)
;Save pricing config scren
;~ 	MouseClick("left",1480,989,1)
;~  	Sleep(3000)

	;following block works on pricing screen
	MouseClick("left",541,241,2)
	Sleep(250)
	Send($sCellValue)
	Sleep(500)
	Send("{ENTER}")
	Sleep(2500)
;~ 	MouseClick("left",56,667,1)
;~ 	Sleep(2500)
	MouseClick("left",355,371,1)
	Sleep(1500)
MouseMove(887,748)
MouseDown("left")
MouseMove(1153,748)
MouseUp("left")

;Change to OLD NTO
;~ MouseClick("left",958,521,1)
;~ Sleep(200)
;~ Send("old nto{ENTER}")
;~ Sleep(200)
;~ MouseClick("left",957,571,1)
;~ Sleep(200)
;~ Send("old nto{ENTER}")
;~ Sleep(200)
;~ MouseClick("left",948,615,1)
;~ Sleep(200)
;~ Send("old nto{ENTER}")
;~ Sleep(200)
;~ MouseClick("left",962,670,1)
;~ Sleep(200)
;~ Send("old nto{ENTER}")
;~ Sleep(200)
;~ MouseClick("left",972,712,1)
;~ Sleep(200)
;~ Send("old nto{ENTER}")
;~ Sleep(200)
;~ MouseMove(1317,589)
;~ MouseDown("left")
;~ MouseMove(1315,681)
;~ MouseUp("left")
;~ Sleep(200)
;~ MouseClick("left",952,721,1)
;~ Sleep(200)
;~ Send("old nto{ENTER}")

;Change to all customer
MouseClick("left",958,521,1)
Sleep(200)
Send("all{ENTER}")
Sleep(200)
MouseClick("left",957,571,1)
Sleep(200)
Send("all{ENTER}")
Sleep(200)
MouseClick("left",948,615,1)
Sleep(200)
Send("all{ENTER}")
Sleep(200)
MouseClick("left",962,670,1)
Sleep(200)
Send("all{ENTER}")
Sleep(200)
MouseClick("left",972,712,1)
Sleep(200)
Send("all{ENTER}")
Sleep(200)
MouseMove(1317,589)
MouseDown("left")
MouseMove(1315,681)
MouseUp("left")
Sleep(200)
MouseClick("left",952,721,1)
Sleep(200)
Send("all{ENTER}")

Sleep(200)
MouseClick("left",1715,838,1)
Sleep(2500)
;~ MouseClick("left",66,638,1)
;~ Sleep(2500)
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
