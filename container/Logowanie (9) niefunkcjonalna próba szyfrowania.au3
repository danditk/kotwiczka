Global $Program_name = "Tworzenie kotwiczki z Darkiem :D"
#include <Array.au3>
Loging()
Func Loging()

	Global $Pass[10][10] = [[0, 1, 2, 3, 4, 5, 6, 7, 8, 9], _
			[@UserName, 'log_in', 'pass_in', 'log_set', 'pass_set', 'state', 'pass', 'gli_loging', 'password'], _
			['dtkaczuk', 'Stasiu', 'PCy', 'glitkaczda', 'glinoconto', 'glikozowpa', 'gliwoliclu'], _
			['admin', 'no', 123, 'glitkaczda', 'glinoconto', 'glikozowpa', 'gliwoliclu'], _
			['paprotka' & @MDAY & @MON, ' ', 123, 'Walmedda', 'Walmedto', 'Walmedpa', 'Walmedlu'], _
			['Admin Darek', 'Darek', 'Tomku', 'Darek', 'Tomaszu', 'Pawel', 'Lukasz'], _
			['Welcome', 'Wrong'], _
			['v', 'x', 'y', 'z', 'i', 'j', 'k', 'a', 'b', 'c']]
	$Pass[1][3] = $Pass[1][7] ; wyswietlanie w input login = gli_login
	$Pass[1][4] = $Pass[1][8] ; wyswietlanie w input haslo = password


	#include <ButtonConstants.au3>
	#include <EditConstants.au3>
	#include <GUIConstantsEx.au3>
	#include <StaticConstants.au3>
	#include <WindowsConstants.au3>
	#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_tworzenie_kotwiczk_loging_1.1.kxf
	Local $Okno_Tworzenie_kotwiczki = GUICreate($Program_name, 411, 179, 219, 126)
	GUISetFont(8, 400, 0, "Showcard Gothic")
	GUISetBkColor(0x313131)
	Local $Button_Enter = GUICtrlCreateButton("Enter", 216, 136, 179, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Local $Pic1 = GUICtrlCreatePic("C:\Users\glitkaczda\Desktop\Programowanie\Program Darka do tworzenia kotwiczek\Valmet_picture.jpg", 216, 16, 180, 105)
	Local $Text_login = GUICtrlCreateLabel("Login", 16, 16, 94, 34, BitOR($SS_CENTER, $SS_NOPREFIX))
	GUICtrlSetFont(-1, 18, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	Local $Text_password = GUICtrlCreateLabel("Password", 16, 96, 158, 34, BitOR($SS_CENTER, $SS_NOPREFIX))
	GUICtrlSetFont(-1, 18, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	$Pass[1][3] = GUICtrlCreateInput($Pass[1][3], 16, 56, 185, 28, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER, $ES_LOWERCASE)) ; gli_login = log_in_obiekt
	GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetCursor(-1, 5)
	$Pass[1][4] = GUICtrlCreateInput($Pass[1][4], 16, 136, 185, 28, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER, $ES_PASSWORD)) ; password = pass_in_obiekt
	GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetCursor(-1, 5)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	While 1

		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit
			Case $Button_Enter

				$v = $Pass[0][0] ; Przypisanie v = 0 za pomoca zmiennych tablicowych Pass
				$x = $Pass[$v][1] ; Przypisanie x = 1 za pomoca zmiennych tablicowych Pass
				$y = $Pass[$v][2] ; Przypisanie y = 2 za pomoca zmiennych tablicowych Pass
				$z = $Pass[$v][3] ; Przypisanie z = 3 za pomoca zmiennych tablicowych Pass
				$i = $Pass[$v][4] ; Przypisanie i = 4 za pomoca zmiennych tablicowych Pass
				$j = $Pass[$v][5] ; Przypisanie j = 5 za pomoca zmiennych tablicowych Pass
				$k = $Pass[$v][6] ; Przypisanie k = 6 za pomoca zmiennych tablicowych Pass
				$a = $Pass[$v][7] ; Przypisanie a = 7 za pomoca zmiennych tablicowych Pass
				$b = $Pass[$v][8] ; Przypisanie b = 8 za pomoca zmiennych tablicowych Pass
				$c = $Pass[$v][9] ; Przypisanie c = 9 za pomoca zmiennych tablicowych Pass

				$Pass[$x][$x] = GUICtrlRead($Pass[$x][$z]) ; log_in = wartosc log_in
				$Pass[$x][$y] = GUICtrlRead($Pass[$x][$i]) ; log_in = wartosc log_in

				Switch $Pass[$x][$x] ; warunek wartosc log_in
					Case $Pass[$z][$v] ; wartosc log_in = administrator
						If $Pass[$x][$y] == $Pass[$i][$v] Then ; warunek pass_in -> login ok -> haslo ok -> odpal program
							MsgBox($v, $Program_name, $Pass[$k][$v] & ' ' & $Pass[$j][$v] & ' !')
						Else
							MsgBox($v, $Program_name, $Pass[$k][$x] & ' ' & $Pass[$x][$a] & ' or ' & $Pass[$x][$k])
						EndIf

					Case $Pass[$z][$x] ; wartosc log_in = darek
						Local $i, $Pass_New = $Pass
						If $Pass[$x][$y] == $Pass[$i][$x] Then ; warunek pass_in = haslo ok
							$Pass_New = _ArrayExtract($Pass, $y, $y, $v, $z) ; tworzenie tablicy PC, size $x z $y by array
							If(_ArraySearch($Pass_New, $Pass[$x][$v])) > -$x Then ; PC zgodny z tablica -> login ok -> haslo ok -> PC ok -> odpal program
								MsgBox($v, $Program_name, 'test no - pc ok')
							Else
								MsgBox($v, $Program_name, 'test no - pc NOT ok')
							EndIf
						Else
							MsgBox($v, $Program_name, 'test no - zle haslo')
						EndIf

					Case $Pass[$z][$y] ; wartosc log_in = tomek / sprawdzenie tez dla mojego PC -  Or $Pass[$x][$v] = $Pass[$y][$z] /
						If($Pass[$x][$y] == $Pass[$i][$y] Or $Pass[$x][$y] == $Pass[$i][$i]) And($Pass[$x][$v] = $Pass[$y][$i] Or $Pass[$x][$v] = $Pass[$y][$z]) Then MsgBox($v, $Program_name, 'test $x$y$z - pc ok') ;  -> login ok -> haslo ok -> PC ok -> odpal program

					Case $Pass[$z][$z], $Pass[$z][$i], $Pass[$z][$j], $Pass[$z][$k] ; wartosc log_in = urzytkownik valmet / zawezona grupa /
						Local $i, $Pass_New = $Pass
						_ArrayTranspose($Pass_New)
						$i = _ArraySearch($Pass_New, $Pass[$x][$x], $v, $v, $v, $v, $x, $z)
						If $Pass[$x][$y] == $Pass[$i][$i] Then
							If $Pass[$x][$v] = $Pass[$y][$i] Then
								MsgBox($v, $Program_name, 'PC: ' & $Pass[$x][$v] & @CRLF & 'Login: ' & $Pass[$x][$x] & @CRLF & 'Haslo: ' & $Pass[$x][$y] & @CRLF & 'Index: ' & $i)
							EndIf
						EndIf
					Case Else
						MsgBox($v, $Program_name, 'Wrong, Case Else' & @CR & 'Login: ' & $Pass[$x][$x] & @CR & 'Haslo: ' & $Pass[$x][$y])
				EndSwitch
		EndSwitch
		If $nMsg = $Button_Enter Then ; Nadpisywanie loginu i hasla przed ponowna petla
			GUICtrlSetData($Pass[$x][$z], $Pass[$x][$a])
			GUICtrlSetData($Pass[$x][$i], $Pass[$x][$b])
		EndIf

	WEnd

EndFunc   ;==>Loging
