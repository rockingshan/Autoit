#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
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
;$oExcel = _ExcelBookOpen(@ScriptDir & '\CHAYANDA_DETAILS.xls', 1, 0)
;_ExcelSheetActivate($oExcel, "Sheet1" )
_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
Local $oExcel_1 = _ExcelBookOpen(@ScriptDir & '\CHAYANDA.xls', 0, 0)
$i = 219
While $i < 300
   $j =$i
    $sCellValue = _ExcelReadCell($oExcel_1, $j, 1)
MouseClick("left",573,202,1)
Send($sCellValue)
Send("{ENTER}")
Sleep(3000)
MouseClick("left",584,355,1)
MouseClick("left",582,498,1)
Send("{CTRLDOWN}c{CTRLUP}")
$content = ClipGet()
$array = StringSplit($content, @LF)
_WinWaitActivate("TEST.txt - Notepad","")
Send($sCellValue & "{TAB}")
Send($content)
_ClipBoard_Empty()
;_WinWaitActivate("Microsoft Excel - CHAYANDA_DETAILS.xls  [Compatibility Mode]","")
;_ExcelWriteCell($oExcel, $sCellValue, $i, 1)
;For $j = 1 To $array[0]
 ;  $k = $j + 1
;   MsgBox(0, '', $array[$j])
;   _ExcelWriteArray($oExcel, $i, $k, $array[$i])
;	  Next

_WinWaitActivate("Meghbela Digital SMS - Mozilla Firefox","")
	$i = $i +1
   WEnd
;   _ExcelBookSaveAs($oExcel, @ScriptDir & "\CHAYANDA_DETAILS.xls", "xls", 0, 1) ; Now we save it into the temp directory; overwrite existing file if necessary
;_ExcelBookClose($oExcel)
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
