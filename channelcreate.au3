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

_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\ALLCHANNEL.xls', 0, 0)

$i = 320
While $i < 327
   $j = $i
   $sCellValue = _ExcelReadCell($oExcel, $j, 1)
   $sCellValue2 = _ExcelReadCell($oExcel, $j, 2)
   MouseClick("left",628,302,1)
   Sleep(250)
   Send($sCellValue)
   Sleep(250)
   MouseClick("left",611,334,1)
   Sleep(250)
   MouseClick("left",555,331,1)
   Sleep(250)
   Send($sCellValue2)
   Sleep(250)
   MouseClick("left",572,357,1)
   Sleep(190)
   MouseClick("left",534,384,1)
   Sleep(190)
   MouseClick("left",539,384,1)
   Sleep(190)
   MouseClick("left",536,432,1)
   Sleep(190)
   MouseClick("left",540,411,1)
   Sleep(190)
   MouseClick("left",535,434,1)
   Sleep(190)
;~    MouseClick("left",535,434,1)
;~    Sleep(250)
;~    Send($sCellValue2)
;~    Sleep(250)
   MouseClick("left",874,384,1)
   Sleep(190)
   MouseClick("left",873,424,1)
   Sleep(190)
   MouseClick("left",881,416,1)
   Sleep(190)
   MouseClick("left",889,432,1)
   Sleep(190)
   MouseClick("left",665,518,1)
   Sleep(190)
   MouseClick("left",652,538,1)
   Sleep(190)
   MouseMove(805,513)
   Sleep(190)
   MouseDown("left")
   MouseMove(856,513)
   MouseUp("left")
   Send("1")
   Sleep(190)
   MouseClick("left",861,515,1)
   Sleep(190)
   MouseClick("left",854,585,1)
   Sleep(190)
   MouseClick("left",460,241,1)
   Sleep(190)
   MouseClick("left",700,317,1)
   Sleep(190)
   MouseClick("left",678,372,1)
   Sleep(190)
   MouseClick("left",429,237,1)
   Sleep(190)
   MouseClick("left",320,187,1)
   Sleep(1500)
   Send("{ENTER}")
   
	$i = $i +1
   WEnd
#endregion --- Au3Recorder generated code End ---

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

