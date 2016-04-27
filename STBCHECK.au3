#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#include <Array.au3>
#region --- Internal functions Au3Recorder Start ---
Global $arr[9999][2]
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
Local $oExcel1 = _ExcelBookOpen(@ScriptDir & '\STB.xls', 0, 0)
ClipPut("")
$k =1
$i = 2
While $i < 3519
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel1, $j, 1)
$arr[$k][0]=$sCellValue
MouseClick("left",587,278,2)
Sleep(150)
Send($sCellValue)
MouseClick("left",676,284,1)
Sleep(190)
MouseMove(954,447)
MouseDown("left")
MouseMove(819,448)
MouseUp("left")
Sleep(180)
Send("{CTRLDOWN}c{CTRLUP}")
$arr[$k][1]=ClipGet()
ClipPut("")
$k = $k + 1


#endregion --- Au3Recorder generated code End ---
$i = $i +1
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

