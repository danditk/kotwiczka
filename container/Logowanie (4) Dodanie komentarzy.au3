$Program_name = "Tworzenie kotwiczki z Darkiem :D"
Loging()
Func Loging()

Global $Pass[10][10] = 	[[0,					1,			2,				3,				4,				5,				6,				7,				8,				9				],	_
						[@UserName,				'log_in',	'pass_in',		'log_set',		'pass_set',		'state',		'pass',			'gli_loging',	'password'						],	_
						['dtkaczuk',			'Stasiu',	'PCy',			'glitkaczda',	'glinoconto',	'glikozowpa',	'gliwoliclu'													],	_
						['admin',				'no',		123,			'glitkaczda',	'glinoconto',	'glikozowpa',	'gliwoliclu'													],	_
						['paprotka'&@MDAY&@MON,	' ',		123,			'Walmed',		'Walmed',		'Walmed',		'Walmed'														],	_
						['Admin Darek',			'Darek',	'Tomku',		'Darek',		'Tomaszu',		'Pawel',		'Lukasz'														],	_
						['Welcome',				'Wrong'																																	]]
$Pass[1][3] = $Pass[1][7] ; wyswietlanie w input login = gli_login
$Pass[1][4] = $Pass[1][8] ; wyswietlanie w input haslo = password

#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_tworzenie_kotwiczk_loging_1.1.kxf
Local $Okno_Tworzenie_kotwiczki = GUICreate("Program_name", 411, 179, 219, 126)
GUISetFont(8, 400, 0, "Showcard Gothic")
GUISetBkColor(0x313131)
Local $Button_Enter = GUICtrlCreateButton("Enter", 216, 136, 179, 33)
GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
GUICtrlSetColor(-1, 0xFFFFFF)
GUICtrlSetBkColor(-1, 0x569557)
GUICtrlSetCursor (-1, 0)
Local $Pic1 = GUICtrlCreatePic("C:\Users\glitkaczda\Desktop\Programowanie\Program Darka do tworzenia kotwiczek\Valmet_picture.jpg", 216, 16, 180, 105)
Local $Text_login = GUICtrlCreateLabel("Login", 16, 16, 94, 34, BitOR($SS_CENTER,$SS_NOPREFIX))
GUICtrlSetFont(-1, 18, 800, 0, "Showcard Gothic")
GUICtrlSetColor(-1, 0x569557)
Local $Text_password = GUICtrlCreateLabel("Password", 16, 96, 158, 34, BitOR($SS_CENTER,$SS_NOPREFIX))
GUICtrlSetFont(-1, 18, 800, 0, "Showcard Gothic")
GUICtrlSetColor(-1, 0x569557)
$Pass[1][3] = GUICtrlCreateInput($Pass[1][3], 16, 56, 185, 28, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER, $ES_LOWERCASE)) ; gli_login = log_in_obiekt
GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
GUICtrlSetColor(-1, 0x569557)
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUICtrlSetCursor (-1, 5)
$Pass[1][4] = GUICtrlCreateInput($Pass[1][4], 16, 136, 185, 28, BitOR($GUI_SS_DEFAULT_INPUT,$ES_CENTER, $ES_PASSWORD)) ; password = pass_in_obiekt
GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
GUICtrlSetColor(-1, 0x569557)
GUICtrlSetBkColor(-1, 0xFFFFFF)
GUICtrlSetCursor (-1, 5)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Exit
		Case $Button_Enter
			$Pass[1][1] = GUICtrlRead($Pass[1][3]) ; log_in = wartosc log_in
			$Pass[1][2] = GUICtrlRead($Pass[1][4]) ; log_in = wartosc log_in
			Switch $Pass[1][1] ; warunek wartosc log_in
				Case $Pass[3][0] ; wartosc log_in = administrator
					$x = 1
					If $Pass[1][2] == $Pass[4][0] Then ; warunek pass_in -> login ok -> haslo ok -> odpal program
						MsgBox(0,$Program_name,$Pass[6][0] & ' ' & $Pass[5][0] & ' !')
					Else
						MsgBox(0,$Program_name,$Pass[6][1] & ' ' & $Pass[1][7] & ' or ' & $Pass[1][6])
					EndIf

				Case $Pass[3][1] ; wartosc log_in = darek
					If $Pass[1][2] == $Pass[4][1] Then ; warunek pass_in = haslo ok
						Local $x, $y1, $y2, $Pass_New[4]  ; tworzenie tablicy PC, size 1 z 2 by for
						$x = 2 ; zakres tablicy w osi x
						$y1 = 0 ; wartosc poczatkowa osi y
						$y2 = $y1 + Ubound($Pass_New) - 1 ; wartosc koncowa osi y
						For $i = $y1 To $y2
							$Pass_New[$i] = $Pass[$x][$i]
						Next
						For $i In $Pass_New ; sprawdzanie czy wartosci tablicy PC = PC
							If $i = $Pass[1][0] Then
								$i = True
								ExitLoop
							Else
								$i = False
							EndIf
						Next
						If $i = True Then ; PC zgodny z tablica -> login ok -> haslo ok -> PC ok -> odpal program
							MsgBox(0,$Program_name,'test no - pc ok')
						Else
							MsgBox(0,$Program_name,'test no - pc NOT ok')
						EndIf
					Else
						MsgBox(0,$Program_name,'test no - zle haslo')
					EndIf

				Case $Pass[3][2] ; wartosc log_in = tomek / sprawdzenie tez dla mojego PC -  Or $Pass[1][0] = $Pass[2][3] /
					If ($Pass[1][2] == $Pass[4][2] Or $Pass[1][2] == $Pass[4][4]) And ($Pass[1][0] = $Pass[2][4] Or $Pass[1][0] = $Pass[2][3]) Then MsgBox(0,$Program_name,'test 123 - pc ok') ;  -> login ok -> haslo ok -> PC ok -> odpal program

				Case $Pass[3][3], $Pass[3][4], $Pass[3][5], $Pass[3][6]  ; wartosc log_in = urzytkownik valmet / zawezona grupa /

					$pc = _ArraySearch($Pass_New, $Pass[1][1],0,0,0,0,1,3)
					MsgBox(0,'','Zalogowano z PC ' & $Pass[1][1])


;~ 					Local $x, $y1, $y2, $Pass_New[7]  ; tworzenie tablicy loginow, size 1 z 2 by for
;~ 						$x = 3 ; zakres tablicy w osi x
;~ 						$y1 = 0 ; wartosc poczatkowa osi y
;~ 						$y2 = $y1 + Ubound($Pass_New) - 1 ; wartosc koncowa osi y
;~ 						For $i = $y1 To $y2
;~ 							$Pass_New[$i] = $Pass[$x][$i]
;~ 						Next
						$i = $Pass[1][1]
;~ 						$pc = _ArraySearch($Pass_New, $Pass[1][1])

;~ 						MsgBox(0,'','Zalogowano z PC ' & $Pass[1][1])


;~ 						For $i = 0 to Ubound($Pass_New) - 1
;~ 							MsgBox(0,'',$Pass_New[$i])
;~ 						Next

				Case Else
					MsgBox(0,$Program_name,'Login: ' & $Pass[1][1] & @CR & 'Haslo: ' & $Pass[1][2])
			EndSwitch
	EndSwitch
	If $nMsg = $Button_Enter Then ; Nadpisywanie loginu i hasla przed ponowna petla
		GUICtrlSetData($Pass[1][3],$Pass[1][7])
		GUICtrlSetData($Pass[1][4],$Pass[1][8])
	EndIf
WEnd
;~ Until $Pass[1][5] = $Pass[1][6]

EndFunc
