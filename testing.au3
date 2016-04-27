#include <MsgBoxConstants.au3>

Example()

Func Example()
    ; Assign a Local variable the number 18.
    Local $iNb = 20
	$i = 0

    ; If the mudulus operation of $iNb by 2 equals to 0 then $iNb is even.
    If Mod($iNb, 10) = 0 Then
		$i = $i + 1
        MsgBox($MB_SYSTEMMODAL, "", $iNb & $i & " is an even number.")
    Else
        MsgBox($MB_SYSTEMMODAL, "", $iNb & $i & " is an odd number.")
    EndIf



EndFunc   ;==>Example
