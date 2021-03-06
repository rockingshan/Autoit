#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#include <Array.au3>
#region --- Internal functions Au3Recorder Start ---
Global $arr[200][6]
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
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel1 = _ExcelBookOpen(@ScriptDir & '\allaccounts.xls', 0, 0)
ClipPut("")
$i = 202
While $i < 204
   $j =$i
   $k =0
    $sCellValue = _ExcelReadCell($oExcel1, $j, 1)
$arr[$k][0]=$sCellValue&"{TAB}"
MouseClick("left",476,212,1)
Sleep(75)
Send($sCellValue)
Sleep(50)
Send("{ENTER}")
Sleep(150)
MouseClick("left",547,324,1)
Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
$arr[$k][1]=ClipGet()&"{TAB}"
ClipPut("")
MouseClick("left",534,354,1)
Sleep(59)
Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
$arr[$k][2]=ClipGet()&"{TAB}"
ClipPut("")
MouseClick("left",528,511,1)
Sleep(75)
Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
$arr[$k][3]=ClipGet()&"{TAB}"
ClipPut("")
MouseClick("left",521,543,1)
Sleep(75)
Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
$arr[$k][4]=ClipGet()&"{TAB}"
ClipPut("")
MouseClick("left",608,569,1)
Sleep(75)
Send("{CTRLDOWN}a{CTRLUP}{CTRLDOWN}c{CTRLUP}")
$arr[$k][5]=ClipGet()
ClipPut("")
;_WinWaitActivate("allcusname.txt - Notepad","")
;$k = 0
;While $k < 6
;   Send($arr[$k])
   $k = $k + 1
;   WEnd
;Send("{ENTER}")
;Sleep(50)
;_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
MouseClick("left",365,166,1)
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

