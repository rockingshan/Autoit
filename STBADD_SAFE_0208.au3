#region ---Au3Recorder generated code Start (v3.3.7.0)  ---
#include <Excel.au3>
#region --- Internal functions Au3Recorder Start ---
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

#AU3RecordSetup()
#endregion --- Internal functions Au3Recorder End ---
Local $oExcel = _ExcelBookOpen(@ScriptDir & '\564.xls', 0, 0)
$i = 2
While $i < 399
_WinWaitActivate("safe view SMS - admin console - lite version — Mozilla Firefox","")
$sCellValue = _ExcelReadCell($oExcel, $i, 1)
MouseClick("left",615,334,2)
Sleep(400)
Send($sCellValue)
Sleep(400)
Send("{TAB}")
$sCellValue = _ExcelReadCell($oExcel, $i, 2)
Sleep(400)
Send($sCellValue)
MouseClick("left",1834,402,1)
Sleep(400)
$i = $i +1
WEnd
#endregion --- Au3Recorder generated code End ---
