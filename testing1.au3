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
#endregion

Local $oExcel = _ExcelBookOpen(@ScriptDir & '\testing.xlsx', 0, 0)
$i = 1
$k = 56
While $i < 4001
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
	If Mod($sCellValue, 100) = 0 Then
		$k = $k + 2
    Else
        $k = $k + 1
    EndIf
	_WinWaitActivate("Untitled - Notepad","")
Send($sCellValue)
Send("{TAB}")
Send($k)
Send("{ENTER}")

#endregion --- Au3Recorder generated code End ---
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



