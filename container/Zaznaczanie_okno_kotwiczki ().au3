
Global $Ustaw_Zmienna = 1
Global $Ustaw_Aktualna = 1
Global $Ustaw_Wl_strony = 1
Zaznaczanie()
Func Zaznaczanie()
Local $Epl_okno_kotwiczki = "Właściwości (symbol graficzny)"
Local $Kategoria = 'ComboBox2'
Local $Kategoria_wartosc = 'Wszystkie kategorie'
Local $Button_zmienna = 2006
Local $Button_aktualna = 2065
Local $Button_wl_strony = 2005
If $Ustaw_Zmienna = 1 Then
	$Status_zmienna = 'Check'
ElseIf $Ustaw_Zmienna = 0 Then
	$Status_zmienna = 'UnCheck'
EndIf
If $Ustaw_Aktualna = 1 Then
	$Status_aktualna = 'Check'
ElseIf $Ustaw_Aktualna = 0 Then
	$Status_aktualna = 'UnCheck'
EndIf
If $Ustaw_Wl_strony = 1 Then
	$Status_wl_strony = 'Check'
ElseIf $Ustaw_Wl_strony = 0 Then
	$Status_wl_strony = 'UnCheck'
EndIf
Do
ControlCommand($Epl_okno_kotwiczki,"",$Kategoria,"SelectString", $Kategoria_wartosc)
ControlCommand($Epl_okno_kotwiczki,"",$Button_zmienna,$Status_zmienna)
ControlCommand($Epl_okno_kotwiczki,"",$Button_aktualna,$Status_aktualna)
ControlCommand($Epl_okno_kotwiczki,"",$Button_wl_strony,$Status_wl_strony)
Local $k1, $k2, $k3, $k4
$ck = ControlCommand($Epl_okno_kotwiczki,"",$Kategoria,"GetCurrentSelection", "")
$cz = ControlCommand($Epl_okno_kotwiczki,"",$Button_zmienna,'IsChecked')
$ca = ControlCommand($Epl_okno_kotwiczki,"",$Button_aktualna,'IsChecked')
$cw = ControlCommand($Epl_okno_kotwiczki,"",$Button_wl_strony,'IsChecked')
Until $ck = $Kategoria_wartosc And $cz = $Ustaw_Zmienna And $ca = $Ustaw_Aktualna And $cw = $Ustaw_Wl_strony
EndFunc