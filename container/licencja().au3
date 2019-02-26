
;~ Licence()
;~ Func Licence()
;~ 	Global $w1 = 20, $w2 = 32
;~ 	$kedy = '404=%5056*788384683045683845634118656704($667&#cr$45=)+56]79]11000110'
;~ 	Local $licence[4] = [@MDAY, StringMid($kedy, $w1, 2), @MON, StringMid($kedy,$w2, 2)]
;~ 	If $licence[0] = $licence[1] And $licence[2] = $licence[3] Then
;~ 		For $l In $licence
;~ 			MsgBox(0, '', $l)
;~ 		Next
;~ 	EndIf
;~ EndFunc   ;==>Licence
$i = 0243
$i = Int($i)

MsgBox(0,'',StringRight(@YEAR,2))