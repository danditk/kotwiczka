#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>
#include <Array.au3>

Global $Loging_Hendle, $Txt_Excel_name, $Excel_full_name, $Sciezka_excela, $Program_Excel_open, $Plik_Excel_open
Global $Epl, $Epl_poz, $licence, $Epl_script, $Epl_okno_kotwiczki_poz, $Epl_okno_kotwiczki, $Ask_ex_open, $Txt_ex_close, $login
Global $Liczba_Meny_how, $licznik, $licznik_x, $Date
$Epl_script = '404=%5056*78838468319568384563406181956704($667&#cr$45=)+56]79]11000110'
$Epl = "EPLAN Electric P8 2.7"
$Txt_Excel_name = 'Kotwiczka'
$Ask_ex_open = 'Czy chcesz otworzyc Excel - ' & $Txt_Excel_name & '?'
$Epl_okno_kotwiczki = "W³aœciwoœci (symbol graficzny)"
$Txt_ex_close = 'Excel - ' & $Txt_Excel_name & ' nie jest otwarty.'
HotKeySet('+!e', 'HotKey_Exit')
Global $Program_name = "Tworzenie kotwiczki z Darkiem :D"

;~ Loging()
;~ GUIDelete($Loging_Hendle)
;~ Licence()
Beep_welcome()
;check state
Excel_Open()
;check state
Program_kotwiczka()

Func Program_kotwiczka()

	#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_okno_tworzenie_kotwiczk_1.2i.kxf
	Global $Program_name = "Tworzenie kotwiczki z Darkiem :D"
	Global $Okno_Tworzenie_kotwiczki = GUICreate("Program_name", 443, 186, 219, 126)
	GUISetFont(8, 400, 0, "Showcard Gothic")
	GUISetBkColor(0x313131)
	Global $Text_excel = GUICtrlCreateLabel("Excel", 336, 16, 94, 34, BitOR($SS_CENTER, $SS_NOPREFIX))
	GUICtrlSetFont(-1, 18, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	Global $Button_Start = GUICtrlCreateButton("Start", 16, 16, 155, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_Start_Many = GUICtrlCreateButton("Start for many", 16, 136, 155, 33, $BS_NOTIFY)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_Excel_Open = GUICtrlCreateButton("Open", 336, 56, 91, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_Excel_Reset = GUICtrlCreateButton("Reset", 335, 96, 91, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_Excel_Close = GUICtrlCreateButton("Close", 336, 136, 91, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Pic1 = GUICtrlCreatePic("C:\Users\glitkaczda\Desktop\Programowanie\Program Darka do tworzenia kotwiczek\Valmet_picture.jpg", 16, 56, 156, 73)
	Global $Meny_how = GUICtrlCreateInput("1", 184, 136, 65, 30, BitOR($GUI_SS_DEFAULT_INPUT, $ES_CENTER))
	GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	GUICtrlSetBkColor(-1, 0xFFFFFF)
	GUICtrlSetCursor(-1, 5)
	Global $Check_Reset = GUICtrlCreateCheckbox("Excel reset", 184, 16, 129, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX, $BS_CENTER, $BS_VCENTER))
	GUICtrlSetState(-1, $GUI_CHECKED)
	GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Check_Zamknij = GUICtrlCreateCheckbox("Exit after", 184, 56, 129, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX, $BS_CENTER, $BS_VCENTER))
	GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Check_Speed = GUICtrlCreateCheckbox("Speed mode", 184, 96, 129, 33, BitOR($GUI_SS_DEFAULT_CHECKBOX, $BS_CENTER, $BS_VCENTER))
	GUICtrlSetFont(-1, 12, 400, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_Up = GUICtrlCreateButton("Up", 260, 137, 51, 17)
	GUICtrlSetFont(-1, 7, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_Down = GUICtrlCreateButton("Down", 260, 152, 51, 17)
	GUICtrlSetFont(-1, 7, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	WinActivate($Program_name)

	While 1

		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Exit_Procedure()

			Case $Button_Start
				Tworzenie_kotwicy()

			Case $Button_Start_Many
				HotKeySet('+d', 'Exit_Procedure')
				$Liczba_Meny_how = GUICtrlRead($Meny_how)
				For $i = 1 To $Liczba_Meny_how
					Tworzenie_kotwicy()
					If $i < $Liczba_Meny_how Then
						WinActivate($Epl)
						AutoItSetOption('MouseCoordMode', 0)
						WinWaitActive($Epl)
						MouseClick('primary', 370, 130, 1, 0)
						AutoItSetOption('MouseCoordMode', 1)
						Send("{PGDN}")
					EndIf
				Next

			Case $Button_Up
				$Liczba_Meny_how = GUICtrlRead($Meny_how)
				$Liczba_Meny_how += 1
				GUICtrlSetData($Meny_how, $Liczba_Meny_how)

			Case $Button_Down
				$Liczba_Meny_how = GUICtrlRead($Meny_how)
				If $Liczba_Meny_how > 1 Then
					$Liczba_Meny_how -= 1
					GUICtrlSetData($Meny_how, $Liczba_Meny_how)
				EndIf

			Case $Button_Excel_Open
				If Not WinExists($Txt_Excel_name) Then
					Excel_Open()
				Else
					MsgBox(0, "Excel info", "Excel jest juz otwarty", 1)
				EndIf

			Case $Button_Excel_Reset
				Excel_Reset()

			Case $Button_Excel_Close
				Excel_Close_Button()

		EndSwitch

	WEnd

EndFunc   ;==>Program_kotwiczka



Func Tworzenie_kotwicy()

	Local $Do_ex_reset, $Handle_Epl_okno_kotwiczki, $Handle_Excel, $Przez_Excel, $Przez_Epl_okno_kotwiczki, $Speed_mode_time
	If GUICtrlRead($Check_Speed) = 1 Then $Speed_mode_time = 0
	If WinExists($Txt_Excel_name) Then
		If WinExists($Epl) Then
			$Epl_poz = WinGetPos($Epl)
			MsgBox(0, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?" & @CRLF & "Jak nic sie nie bedzie dzialo to znaczy, ze to TY cos zrobiles nie tak.", 5)
			$tc = 2000 ;~ tc - copy time
			$ta = 1000 * $Speed_mode_time ;~ ta - approve timen
			$Przez_Epl_okno_kotwiczki = 80
			$Przez_Excel = 80

;~ 1. Skopiowanie nazw pelnych do excela
			WinActivate($Epl)
			AutoItSetOption('MouseCoordMode', 0)
			WinWaitActive($Epl)
			MouseClick('primary', 370, 130, 1, 0)
			AutoItSetOption('MouseCoordMode', 1)
			Sleep(100)
			ControlSend($Epl, '', "[CLASSNN:AfxFrameOrView140u1]", '^{a}')
			Sleep(250)
			ControlSend($Epl, '', 'Afx:0000000140000000:8:0000000000010007:0000000000000010:000000000000000015', '!{t}')
			ControlSend($Epl, '', 'Afx:0000000140000000:8:0000000000010007:0000000000000010:000000000000000015', '{o}')
			WinWait($Epl_okno_kotwiczki, '', 20)
			$Handle_Excel = WinGetHandle($Txt_Excel_name)
			If GUICtrlRead($Check_Speed) = 4 Then
				WinSetTrans($Handle_Excel, "", 0)
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinActivate($Txt_Excel_name)
				For $i = 0 to 255
					WinSetTrans($Handle_Excel, "", $i)
					Sleep(10)
				Next
				WinSetTrans($Handle_Excel, "", 255)
			EndIf
			If GUICtrlRead($Check_Speed) = 1 Then MsgBox(0, $Program_name, 'Kliknij "OK" po ukazaniu sie okienka kotwiczki, by przyspieszyc 3 minutowe odliczanie.', 180)
			WinActivate($Epl_okno_kotwiczki)
			WinWaitActive($Epl_okno_kotwiczki, "", 60)
			If WinActive($Epl_okno_kotwiczki) Then
				Sleep(1000)
				$Epl_okno_kotwiczki_poz = WinGetPos($Epl_okno_kotwiczki)
				WinActivate($Txt_Excel_name)
				If GUICtrlRead($Check_Speed) = 4 Then
					WinMove($Txt_Excel_name, '', $Epl_okno_kotwiczki_poz[0] + 100, $Epl_okno_kotwiczki_poz[1] + 100, $Epl_okno_kotwiczki_poz[2], $Epl_okno_kotwiczki_poz[3], 10)
					For $i = 255 to 0 Step -1
						WinSetTrans($Handle_Excel, "", $i)
						Sleep(10)
					Next
					WinSetTrans($Handle_Excel, "", 0)
				Else
					WinMove($Txt_Excel_name, '', $Epl_okno_kotwiczki_poz[0] + 100, $Epl_okno_kotwiczki_poz[1] + 100, $Epl_okno_kotwiczki_poz[2], $Epl_okno_kotwiczki_poz[3])
				EndIf
				Global $Ustaw_Zmienna = 0
				Global $Ustaw_Aktualna = 1
				Global $Ustaw_Wl_strony = 1
				WinActivate($Epl_okno_kotwiczki)
				Zaznaczanie()
				Sleep(100)
				WinActivate($Epl_okno_kotwiczki)
				$Handle_Epl_okno_kotwiczki = WinGetHandle($Epl_okno_kotwiczki)
				If GUICtrlRead($Check_Speed) = 4 Then
					For $i = 255 to $Przez_Epl_okno_kotwiczki Step -1
						WinSetTrans($Handle_Epl_okno_kotwiczki, "", $i)
						Sleep(20)
					Next
					WinSetTrans($Handle_Epl_okno_kotwiczki, "", $Przez_Epl_okno_kotwiczki)
					WinSetState($Txt_Excel_name, '', @SW_HIDE)
					WinSetTrans($Handle_Excel, "", $Przez_Excel)
				EndIf
				WinActivate($Epl_okno_kotwiczki)
				AutoItSetOption('MouseCoordMode', 0)
				MouseClick('primary', 100, 70, 1, 0)
				Sleep(100)
				MouseClick('primary', 115, 190, 1, 0)
				Sleep(100)
				MouseClick('primary', 115, 190, 1, 0)
				Sleep(100)
				MouseClick('primary', 90, 290, 1, 0)
				AutoItSetOption('MouseCoordMode', 1)
				WinActivate($Epl_okno_kotwiczki)
				WinWaitActive($Epl_okno_kotwiczki)
				Send("!{k}")
				Send("{TAB}")
				Send("^{a}")
				Send("^+{F10}")
				Send("{p}")
				$Ustaw_Zmienna = 1
				$Ustaw_Aktualna = 0
				$Ustaw_Wl_strony = 1
				Zaznaczanie()
				WinActivate($Epl_okno_kotwiczki)
				Send("!{k}")
				Send("{TAB}")
				Send("^{a}")
				Send("^{c}")
				Sleep($tc)

;~ 2. Usuniecie nazw pelnych w excelu
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinActivate($Txt_Excel_name)
				Sleep(300)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{HOME}')
				Sleep(300)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{v}')
				Sleep($tc)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '+{Enter}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{RIGHT}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^+{UP}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '+{DOWN}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{c}')
				Sleep($tc)
				WinSetState($Txt_Excel_name, '', @SW_HIDE)

;~ 3. Umieszczenie nazw wyswietlanych i kopiowanie wlasciwosci
				WinActivate($Epl_okno_kotwiczki)
				WinWaitActive($Epl_okno_kotwiczki)
				Send("{RIGHT 3}")
				Send("{DOWN}")
				Send("^{v}")
				Sleep($tc)
				Send("^{TAB}")
				Sleep(500)
				Send("^{a}")
				Send("^{c}")
				Sleep($tc)

;~ 4. Zamiana Wlasciwosci na prawidlowe w excelu
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinActivate($Txt_Excel_name)
				Sleep(300)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{F5}')
				Sleep(300)
				If GUICtrlRead($Check_Speed) = 4 Then WinSetTrans('Go To', "", $Przez_Excel)
				WinActivate($Txt_Excel_name)
				Sleep(100)
				ControlSend('Go To', '', 'EDTBX1', '{z}')
				Sleep(100)
				ControlSend('Go To', '', 'EDTBX1', '{Enter}')
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{RIGHT}')
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{LEFT}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{v}')
				Sleep($tc)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '+{Enter}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{RIGHT}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^+{UP}')
				Sleep(500)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '+{DOWN}')
				Sleep(100)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{c}')
				Sleep($tc)
				WinSetState($Txt_Excel_name, '', @SW_HIDE)

;~ 5. Stworzenie nowego obiektu wlasciwosci i przypo¿¹dkowanie prawidlowych zmiennych
				WinActivate($Epl_okno_kotwiczki)
				WinWaitActive($Epl_okno_kotwiczki)
				Send("^+{F10}")
				Send("{o}")
				Send("^{v}")
				If GUICtrlRead($Check_Speed) = 4 Then
					For $i = $Przez_Epl_okno_kotwiczki to 120
						WinSetTrans($Handle_Epl_okno_kotwiczki, "", $i)
						Sleep(20)
					Next

;~ 6. Nazwanie kotwiczki
					Send("{TAB}")
					For $i = 120 to 180
						WinSetTrans($Handle_Epl_okno_kotwiczki, "", $i)
						Sleep(20)
					Next
					Send("!{n}")
					Send('PREPLANNING')
					For $i = 180 to 255
						WinSetTrans($Handle_Epl_okno_kotwiczki, "", $i)
						Sleep(20)
					Next
					WinSetTrans($Handle_Epl_okno_kotwiczki, "", 255)
				ElseIf GUICtrlRead($Check_Speed) = 1 Then
					Send("{TAB}")
					Send("!{n}")
					Send('PREPLANNING')
				EndIf
				Sleep($ta)
				Send("{Enter}")
				WinMove($Txt_Excel_name, '', $Epl_poz[0] + 50, $Epl_poz[1] + 100, 0, 0)
				If GUICtrlRead($Check_Speed) = 4 Then
					WinSetTrans($Handle_Excel, "", 255)
				EndIf
				MsgBox(0, $Program_name, "Kotwiczka stworzona ( mam nadzieje ;D )", 3)
				If GUICtrlRead($Check_Zamknij) = 1 Then Exit_Procedure()
				If GUICtrlRead($Check_Reset) = 1 Then
					Excel_Reset()
				Else
					Excel_Do_Reset()
					If $licznik <> $licznik_x Then Excel_Reset()
				EndIf
			Else
				Activate_program_name_err()
			EndIf
		Else
			MsgBox(1, $Program_name, "Program " & $Epl & " nie zostal wlaczony." & @CRLF & "Wlacz " & $Epl & "kliknij na strone i sprobuj ponownie")
		EndIf
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
		$Do_ex_open = MsgBox(1, $Program_name, $Ask_ex_open)
		If $Do_ex_open = 1 Then
			Excel_Open()
		Else
			MsgBox(0, $Program_name, "To moj pierwszy program." & @CRLF & "Do jego poprawnego dzialania potrzebuje tego Excela.")
		EndIf
	EndIf
	WinActivate($Program_name)

EndFunc   ;==>Tworzenie_kotwicy



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

	#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_tworzenie_kotwiczk_loging_1.1.kxf
	Local $Okno_Loging = GUICreate($Program_name & ' LOGOWANIE', 411, 179, 219, 126)
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

	$Loging_Hendle = WinGetHandle($Okno_Loging)

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
							ExitLoop
						EndIf

					Case $Pass[$z][$x] ; wartosc log_in = darek
						Local $i, $Pass_New = $Pass
						If $Pass[$x][$y] == $Pass[$i][$x] Then ; warunek pass_in = haslo ok
							$Pass_New = _ArrayExtract($Pass, $y, $y, $v, $z) ; tworzenie tablicy PC, size $x z $y by array
							If(_ArraySearch($Pass_New, $Pass[$x][$v])) > -$x Then ; PC zgodny z tablica -> login ok -> haslo ok -> PC ok -> odpal program
								ExitLoop
							EndIf
						Else
						EndIf

					Case $Pass[$z][$y] ; wartosc log_in = tomek / sprawdzenie tez dla mojego PC -  Or $Pass[$x][$v] = $Pass[$y][$z] /
						If($Pass[$x][$y] == $Pass[$i][$y] Or $Pass[$x][$y] == $Pass[$i][$i]) And($Pass[$x][$v] = $Pass[$y][$i] Or $Pass[$x][$v] = $Pass[$y][$z]) Then
							ExitLoop ;  -> login ok -> haslo ok -> PC ok -> odpal program
						EndIf

					Case $Pass[$z][$z], $Pass[$z][$i], $Pass[$z][$j], $Pass[$z][$k] ; wartosc log_in = urzytkownik valmet / zawezona grupa /
						Local $i, $Pass_New = $Pass
						_ArrayTranspose($Pass_New)
						$i = _ArraySearch($Pass_New, $Pass[$x][$x], $v, $v, $v, $v, $x, $z)
						If $Pass[$x][$y] == $Pass[$i][$i] Then
							If $Pass[$x][$v] = $Pass[$y][$i] Then
								ExitLoop
							EndIf
						EndIf
					Case Else
				EndSwitch
		EndSwitch
		If $nMsg = $Button_Enter Then ; Nadpisywanie loginu i hasla przed ponowna petla
			GUICtrlSetData($Pass[$x][$z], $Pass[$x][$a])
			GUICtrlSetData($Pass[$x][$i], $Pass[$x][$b])
		EndIf

	WEnd

EndFunc   ;==>Loging



Func Excel_Open()

	Global $licznik = 1, $licznik_x = 3
	If Not WinExists($Txt_Excel_name) Then
		Local $ok = MsgBox(1, $Program_name, "Poczekaj, az odpali sie excel. Ok?", 5)
		If $ok <> 7 And $ok <> 2 Then
			WinActivate($Txt_Excel_name)
			Local $ukosnik, $rozszerzenie
			$ukosnik = '\'
			$rozszerzenie = '.xlsx'
			$Excel_full_name = $Txt_Excel_name & $rozszerzenie
			$ukosnik &= $Excel_full_name
			$Sciezka_excela = @ScriptDir & $ukosnik
			$Program_Excel_open = _Excel_Open()
			$Plik_Excel_open = _Excel_BookOpen($Program_Excel_open, $Sciezka_excela)
			WinWait($Txt_Excel_name, "", 15)
			If WinExists($Txt_Excel_name) Then
				If WinExists($Epl) Then
					WinMove($Txt_Excel_name, '', $Epl_poz[0] + 50, $Epl_poz[1] + 100, 0, 0)
				Else
					MsgBox(0, $Program_name, "Program " & $Epl & " nie zostal wlaczony." & @CRLF & "Pamietaj by wlaczyc " & $Epl, 3)
				EndIf
				WinSetState($Txt_Excel_name, '', @SW_HIDE)
				MsgBox(0, "Excel info", "Excel jest gotowy do dzialania", 3)
				WinActivate($Program_name)
			Else
				Activate_program_name_err()
			EndIf
		EndIf
	Else
		Excel_Do_Reset()
		If $licznik <> $licznik_x Then Excel_Reset()
	EndIf

EndFunc   ;==>Excel_Open

Func Excel_Close()

	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		WinWaitActive($Txt_Excel_name, "", 20)
		WinClose($Txt_Excel_name)
		Sleep(500)
		Send("{n}")
		MsgBox(0, "Excel", "Excel zostal zamkniety", 1)
		WinActivate($Program_name)
	EndIf

EndFunc   ;==>Excel_Close

Func Excel_Close_Button()

	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		WinClose($Txt_Excel_name)
		Sleep(500)
		Send("{n}")
		MsgBox(0, "Excel", "Excel zostal zamkniety", 1)
		WinActivate($Program_name)
	Else
		MsgBox(0, $Program_name, $Txt_ex_close, 3)
	EndIf

EndFunc   ;==>Excel_Close_Button

Func Excel_Do_Reset()

	Do
		$Do_ex_reset = MsgBox(4, $Program_name, "Czy chcesz zresetowac excela?", 8)
		If $Do_ex_reset = $IDNO Then
			$Do_ex_show = MsgBox(4, $Program_name, "Chcesz zobaczyc dane z Excel - " & $Txt_Excel_name & "?", 10)
			If $Do_ex_show = 6 Then
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinMove($Txt_Excel_name, '', $Epl_poz[0] + 200, $Epl_poz[1] + 100, 1700, 800)
				WinActivate($Txt_Excel_name)
				Do
					Sleep(10000)
					$ok = MsgBox(4, $Program_name, "Napatrzyles sie juz?")
				Until $ok = 6
				WinMove($Txt_Excel_name, '', $Epl_poz[0] + 50, $Epl_poz[1] + 100, 0, 0)
				WinSetState($Txt_Excel_name, '', @SW_HIDE)
			EndIf
		EndIf
		$licznik += 1
	Until $Do_ex_reset <> $IDNO Or $licznik = $licznik_x
	If $licznik = $licznik_x Then
		$Do_ex_show = MsgBox(4, $Program_name, "Chcesz zobaczyc dane z Excel" & $Txt_Excel_name, 8)
		If $Do_ex_show = 6 Then
			WinSetState($Txt_Excel_name, '', @SW_SHOW)
			WinMove($Txt_Excel_name, '', 100, 100, 1700, 800)
			WinActivate($Txt_Excel_name)
		EndIf
	EndIf

EndFunc   ;==>Excel_Do_Reset

Func Excel_Reset()

	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{HOME}')
		Sleep(100)
		For $i = 1 To 10
			ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{z}')
			Sleep(100)
		Next
		Sleep(100)
		ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{HOME}')
		Sleep(100)
		If GUICtrlRead($Check_Reset) = 4 Then MsgBox(0, "Excel info", "Excel czysty", 2)
		WinActivate($Program_name)
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
		$Do_ex_open = MsgBox(1, $Program_name, $Ask_ex_open)
		If $Do_ex_open = 1 Then
			Excel_Open()
		EndIf

	EndIf

EndFunc   ;==>Excel_Reset



Func Zaznaczanie()

	Local $Kategoria = 'ComboBox2'
	Local $Kategoria_wartosc = 'Wszystkie kategorie'
	Local $Button_Zmienna = 2006
	Local $Button_Aktualna = 2065
	Local $Button_Wl_strony = 2005
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
		ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "SelectString", $Kategoria_wartosc)
		Sleep(2000)
		Local $k1
		$k1 = ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "GetCurrentSelection", "")
	Until $k1 = $Kategoria_wartosc

	Do
		ControlCommand($Epl_okno_kotwiczki, "", $Button_Zmienna, $Status_zmienna)
		ControlCommand($Epl_okno_kotwiczki, "", $Button_Aktualna, $Status_aktualna)
		ControlCommand($Epl_okno_kotwiczki, "", $Button_Wl_strony, $Status_wl_strony)
		Local $k2, $k3, $k4
		$k2 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Zmienna, 'IsChecked')
		$k3 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Aktualna, 'IsChecked')
		$k4 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Wl_strony, 'IsChecked')
		Sleep(2000)
	Until $k2 = $Ustaw_Zmienna And $k3 = $Ustaw_Aktualna And $k4 = $Ustaw_Wl_strony
EndFunc   ;==>Zaznaczanie

Func Licence()
	Global $licence[6] = [@MDAY, StringMid($Epl_script, 20, 2), @MON, StringMid($Epl_script, 32, 2), @YEAR, StringMid($Epl_script, 36, 2)]
	If $licence[0] <= $licence[1] And $licence[2] <= $licence[3] And Int(StringRight($licence[4], 2)) <= $licence[5] Then
		Date()
		$i = int($licence[3])
		$licence[3] = $Date[$i]
		MsgBox(0, $Program_name, 'The license is current to: ' & $licence[1] & " " & $licence[3] & " " & $licence[5], 2)
		MsgBox(0, '', $i)
	Else
		MsgBox(0, $Program_name, 'The license is not current', 8)
		Exit_Procedure()
	EndIf

EndFunc   ;==>Licence

Func Activate_program_name_err()

	MsgBox(0, $Program_name, "Chyba cos poszlo nie tak", 3)
	WinActivate($Program_name)

EndFunc   ;==>Activate_program_name_err


Func HotKey_Exit()
	Exit
EndFunc   ;==>HotKey_Exit

Func Exit_Procedure()
	Excel_Close()
	Beep_bye()
	MsgBox(0, $Program_name, "Wszystko pozamykane, dziekuje " & @CRLF & @CRLF & "BYE!", 3)
	Exit
EndFunc   ;==>Exit_Procedure

Func Beep_welcome()
	For $i = 400 to 1000 Step 250
		Beep($i, 100)
	Next
	Sleep(100)
	Beep(2000, 800)
EndFunc   ;==>Beep_welcome

Func Beep_bye()
	For $i = 2000 to 800 Step -250
		Beep($i, 100)
	Next
	Sleep(100)
	Beep(400, 800)
EndFunc   ;==>Beep_bye


Func Variables()
EndFunc   ;==>Variables

Func Date()
	Global $Date[13] = _
			[0, 'January', 'February', 'March', 'April', _
			'May', 'June', 'July', 'August', _
			'September', 'October', 'November', 'December']
EndFunc   ;==>Date