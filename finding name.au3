#region --- Au3Recorder generated code Start (v3.3.9.5 KeyboardLayout=00000409)  ---
#include <Excel.au3>
#include <Clipboard.au3>
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
#endregion --- Internal functions Au3Recorder End ---

_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\only account.xlsx', 0, 0)
$i = 50
While $i < 236
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel, $j, 1)
MouseClick("left",494,197,1)
Sleep(150)
Send($sCellValue)
Sleep(150)
Send("{ENTER}")
Sleep(150)
MouseMove(658,323)
MouseDown("left")
MouseMove(315,320)
MouseUp("left")
Sleep(190)
Send("{CTRLDOWN}c{CTRLUP}")
$content = ClipGet()
$array = StringSplit($content, @LF)
_WinWaitActivate("name search.txt - Notepad","")
Send($sCellValue & "{TAB}")
Send($content & "{ENTER}")
_ClipBoard_Empty()
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
	$i = $i +1
   WEnd
#endregion --- Au3Recorder generated code End ---
While 1
    Sleep(100)
WEnd

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