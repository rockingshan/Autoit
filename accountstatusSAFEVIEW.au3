#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#include <Array.au3>
#region --- Internal functions Au3Recorder Start ---
Global $arr[9999][3]
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
Local $oExcel2 = _ExcelBookNew()
_WinWaitActivate("Safe View - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\NEXUS.xls', 0, 0)
$i = 2
$k =1
While $i < 135
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	$arr[$k][0]=$sCellValue
MouseClick("left",251,203,1)
Sleep(75)
MouseClick("left",303,320,1)
Sleep(150)
Send($sCellValue)
Send("{ENTER}")
Sleep(130)
MouseClick("left",185,342,1)
;~ MouseClick("left",594,329,2)
;~ Send("{CTRLDOWN}c{CTRLUP}")
;~ $arr[$k][1]=ClipGet()
;~ Sleep(130)
;~ _WinWaitActivate("Untitled - Notepad","")
;~ Send($sCellValue)
;~ Send("{TAB}")
;~ Send("{CTRLDOWN}v{CTRLUP}")
;~ _WinWaitActivate("DuoSubscribe Version 5.0 - Mozilla Firefox","")
;~ MouseClick("left",530,267,1)
;~ Sleep(120)
MouseMove(567,362)
MouseDown("left")
MouseMove(879,359)
MouseUp("left")
Send("{CTRLDOWN}c{CTRLUP}")
$arr[$k][1]=ClipGet()
Sleep(110)
;~ _WinWaitActivate("Untitled - Notepad","")
;~ Send("{TAB}")
;~ Send("{CTRLDOWN}v{CTRLUP}")
;~ Send("{ENTER}")
;~ Sleep(150)
;~ _WinWaitActivate("DuoSubscribe Version 5.0 - Mozilla Firefox","")
;~ MouseMove(373,174)
;~ MouseDown("left")
;~ MouseMove(374,174)
;~ MouseUp("left")
ClipPut("")
#endregion --- Au3Recorder generated code End ---
$i = $i +1
$k = $k + 1
   WEnd
_ExcelWriteSheetFromArray($oExcel2, $arr,1,1,0,0)
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

