#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Compression=4
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
Func RemoveLeadingZero($num)
    While StringLeft($num, 1) = "0"
        $num = StringTrimLeft($num, 1)
    WEnd
    Return $num
EndFunc
Func _WinWaitActivate($title,$text,$timeout=0)
	WinWait($title,$text,$timeout)
	If Not WinActive($title,$text) Then WinActivate($title,$text)
	WinWaitActive($title,$text,$timeout)
EndFunc

; Get the current day and month
$day = RemoveLeadingZero(@MDAY)
$month = RemoveLeadingZero(@MON)

; Create the filename
$filename = "Elogbooks_" & $day & "_" & $month & ".pdf"
;create a function
Func WhatsAppSender($text)
	;open whatsapp. whatsapp need to be installed
	$command = "whatsapp://send?phone=" & $text & "&text=Daily school log"
	;ShellExecute("whatsapp://send?phone=918420008648&text=Daily school log")
	ShellExecute($command)
	_WinWaitActivate("WhatsApp","")
	WinSetState("WhatsApp", "", @SW_MAXIMIZE)
	Sleep(600)
	MouseClick("left",500,1022,1)
	Sleep(250)
	MouseUp("left")
	Sleep(250)
	MouseClick("left",516,887,1)
	Sleep(250)
	MouseUp("left")
	Sleep(1500)
	_WinWaitActivate("Open","")
	Sleep(250)
	MouseClick("left",452,47,1)
	Sleep(250)
	Send("C:\Users\Shantanu\Downloads")
	Sleep(250)
	Send("{ENTER}")
	Sleep(1000)
	MouseClick("left",665,477,1)
	Sleep(250)
	Send($filename)
	MouseClick("left",773,504,1)
	Sleep(3000)
	_WinWaitActivate("WhatsApp","")
	Sleep(250)
	MouseClick("left",983,1013,1)
	Sleep(250)
	MouseUp("left")
	Sleep(250)
EndFunc
WhatsAppSender(918420008648)
WhatsAppSender(916296734717)
WinMinimizeAll()
;Send("{LWIN}d")