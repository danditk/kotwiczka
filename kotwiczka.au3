#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>
#include <Array.au3>

Func Initiation()

;~ 1. Zmienne Main GUI
	Global $Program_name, $Okno_Tworzenie_kotwiczki, $Text_exce, $Button_Start, $Button_Start_Many, $nMsg
	Global $Button_Excel_Open, $Button_Excel_Reset, $Button_Excel_Close, $Pic1, $Meny_how, $Check_Reset, $Check_Zamknij
	Global $Check_Speed, $Button_Up, $Button_Down
	$Program_name = "Tworzenie kotwiczki z Darkiem :D"

;~ 2. Zmienne MSGBox_ok GUI
	Global $Handle_MSGBox_ok = WinGetHandle($Program_name & ' INFO')

;~ 3. Zmienne MSGBox_KO GUI
	Global $Handle_MSGBox_KO = WinGetHandle($Program_name & ' INFO')

;~ 4. Zmienne MSGBox_NOYES GUI
	Global $Handle_MSGBox_yesno = WinGetHandle($Program_name & ' ASK')

;~ 5. Zmienne MSGBox_NOYES GUI
	Global $Handle_MSGBox_NOYES = WinGetHandle($Program_name & ' ASK')

	Global $Loging_Hendle, $Txt_Excel_name, $Excel_full_name, $Sciezka_excela, $Program_Excel_open, $Plik_Excel_open
	Global $Epl, $Epl_poz, $licence, $Epl_script, $Epl_okno_kotwiczki_poz, $Epl_okno_kotwiczki, $Ask_ex_open, $Txt_ex_close, $login
	Global $Liczba_Meny_how, $i_many, $licznik, $licznik_x, $Date, $Handle_Excel
	Global $Ustaw_Zmienna, $Ustaw_Aktualna, $Ustaw_Wl_strony

	$Epl_script = '404=%5056*78838468319568384563406181956704($667&#cr$45=)+56]79]11000110'
	$Epl = 'EPLAN'
	$Txt_Excel_name = 'Kotwiczka'
	$Ask_ex_open = 'Czy chcesz otworzyc Excel - ' & $Txt_Excel_name & '?'
	$Epl_okno_kotwiczki = "W³aœciwoœci (symbol graficzny)"
	$Txt_ex_close = 'Excel - ' & $Txt_Excel_name & ' nie jest otwarty.'
	$Epl_poz = WinGetPos($Epl)
	HotKeySet('+!e', 'HotKey_Exit')
	Global $Program_name = "Tworzenie kotwiczki z Darkiem :D"

EndFunc   ;==>Initiation

Initiation()
Beep_welcome()
Loging()
GUIDelete($Loging_Hendle)
Licence()
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
				Tworzenie_kotwicy_many()

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
	$Liczba_Meny_how = GUICtrlRead($Meny_how)
	If GUICtrlRead($Check_Speed) = 1 Then $Speed_mode_time = 0
	If WinExists($Txt_Excel_name) Then
		If WinExists($Epl) Then
			$Epl_poz = WinGetPos($Epl)
			MsgBox(1, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?" & @CRLF & "Jak nic sie nie bedzie dzialo to znaczy, ze to TY cos zrobiles nie tak.", 5)
			$tc = 2000 ;~ tc - copy time
			$ta = 1000 * $Speed_mode_time ;~ ta - approve timen
			$Przez_Epl_okno_kotwiczki = 80
			$Przez_Excel = 80

			;1. Skopiowanie nazw pelnych do excela
			WinActivate($Epl)
			AutoItSetOption('MouseCoordMode', 0)
			WinWaitActive($Epl)
			MouseClick('primary', 1300, 500, 1, 0)
			AutoItSetOption('MouseCoordMode', 1)
			Sleep(500)
			ControlSend($Epl, '', "[CLASSNN:AfxFrameOrView140u1]", '^{a}')
			Sleep(500)
			ControlSend($Epl, '', 'Afx:0000000140000000:8:0000000000010003:0000000000000010:000000000000000015', '!{t}')
			ControlSend($Epl, '', 'Afx:0000000140000000:8:0000000000010003:0000000000000010:000000000000000015', '{o}')
			If GUICtrlRead($Check_Speed) = 1 Then MsgBox(0, $Program_name, 'Kliknij "OK" po ukazaniu sie okienka kotwiczki, by przyspieszyc 3 minutowe odliczanie.', 180)
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
			WinActivate($Epl_okno_kotwiczki)
			WinWaitActive($Epl_okno_kotwiczki, "", 60)
			If WinActive($Epl_okno_kotwiczki) Then
				If GUICtrlRead($Check_Speed) = 0 Then Sleep(10000)
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
				WinActivate($Epl_okno_kotwiczki)
				Zaznaczanie()
				Sleep(100)
				WinActivate($Epl_okno_kotwiczki)
				$Handle_Epl_okno_kotwiczki = WinGetHandle($Epl_okno_kotwiczki)
				If GUICtrlRead($Check_Speed) = 4 Then
					For $i = 255 to $Przez_Epl_okno_kotwiczki Step -1
						WinSetTrans($Handle_Epl_okno_kotwiczki, "n", $i)
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
				Zaznaczanie(1, 0, 1)
				WinActivate($Epl_okno_kotwiczki)
				Send("!{k}")
				Send("{TAB}")
				Send("^{a}")
				Send("^{c}")
				Sleep($tc)

				;2. Usuniecie nazw pelnych w excelu

				$Epl_okno_kotwiczki_poz = WinGetPos($Epl_okno_kotwiczki)
				WinMove($Txt_Excel_name, '', $Epl_okno_kotwiczki_poz[0] + 100, $Epl_okno_kotwiczki_poz[1] + 100, $Epl_okno_kotwiczki_poz[2], $Epl_okno_kotwiczki_poz[3], 10)
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

				;3. Umieszczenie nazw wyswietlanych i kopiowanie wlasciwosci
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

				;4. Zamiana Wlasciwosci na prawidlowe w excelu
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinActivate($Txt_Excel_name)
				Sleep(300)
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{F5}')
				Sleep(300)
				WinSetTrans('Go To', "", $Przez_Excel)
				WinActivate($Txt_Excel_name)
				Sleep(100)
				ControlSend('Go To', '', 'EDTBX1', '{z}')
				Sleep(100)
				ControlSend('Go To', '', 'EDTBX1', '{Enter}')
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{RIGHT}')
				Sleep(100)
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

				;5. Stworzenie nowego obiektu wlasciwosci i przypo¿¹dkowanie prawidlowych zmiennych
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

					;6. Nazwanie kotwiczki
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
				If $nMsg = $Button_Start Then
					MsgBox(0, $Program_name, "Kotwiczka stworzona ( mam nadzieje ;D )", 3)
				Else
					Local $literki = 'ki'
					If $Liczba_Meny_how >= 5 Then $literki = 'ek'
					MsgBox(0, $Program_name, $Liczba_Meny_how & ' Twoich kotwicz' & $literki & ' zostalo utworzone', 3)
				EndIf
				If(GUICtrlRead($Check_Zamknij) = 1 And $nMsg = $Button_Start) Or(GUICtrlRead($Check_Zamknij) = 1 And $nMsg = $Button_Start_Many And $i_many >= $Liczba_Meny_how) Then Exit_Procedure()
				If GUICtrlRead($Check_Reset) = 1 Then
					Excel_Reset()
				Else
					Excel_Do_Reset()
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
				Exit_Procedure()
			Case $Button_Enter
				$Pass[1][1] = GUICtrlRead($Pass[1][3]) ; log_in = wartosc log_in
				$Pass[1][2] = GUICtrlRead($Pass[1][4]) ; log_in = wartosc log_in
				Switch $Pass[1][1] ; warunek wartosc log_in
					Case $Pass[3][0] ; wartosc log_in = administrator
						$x = 1
						If $Pass[1][2] == $Pass[4][0] Then ; warunek pass_in -> login ok -> haslo ok -> odpal program
;~ 							MsgBox(0, $Program_name, $Pass[6][0] & ' ' & $Pass[5][0] & ' !') ; temporary
							ExitLoop
						Else
;~ 							MsgBox(0, $Program_name, $Pass[6][1] & ' ' & $Pass[1][7] & ' or ' & $Pass[1][6]) ; temporary
						EndIf

					Case $Pass[3][1] ; wartosc log_in = darek
						If $Pass[1][2] == $Pass[4][1] Then ; warunek pass_in = haslo ok

							$Pass_New = _ArrayExtract($Pass, 2, 2, 0, 3) ; tworzenie tablicy PC, size 1 z 2 by array
;~ 							_ArrayDisplay($Pass_New)
							$i = _ArraySearch($Pass_New, $Pass[1][0])
							ExitLoop

;~ 							Local $x, $y1, $y2, $Pass_New[4] ; tworzenie tablicy PC, size 1 z 2 by for
;~ 							$x = 2 ; zakres tablicy w osi x
;~ 							$y1 = 0 ; wartosc poczatkowa osi y
;~ 							$y2 = $y1 + Ubound($Pass_New) - 1 ; wartosc koncowa osi y
;~ 							For $i = $y1 To $y2
;~ 								$Pass_New[$i] = $Pass[$x][$i]
;~ 							Next

;~ 							For $i In $Pass_New ; sprawdzanie czy wartosci tablicy PC = PC
;~ 								If $i = $Pass[1][0] Then
;~ 									$i = True
;~ 									ExitLoop
;~ 								Else
;~ 									$i = False
;~ 								EndIf
;~ 							Next

;~ 							If $i = True Then ; PC zgodny z tablica -> login ok -> haslo ok -> PC ok -> odpal program

;~ 							If $i > -1 Then
;~ 								MsgBox(0, $Program_name, 'test no - pc ok')
;~ 							Else
;~ 								MsgBox(0, $Program_name, 'test no - pc NOT ok')
;~ 							EndIf
						Else
							MsgBox(0, $Program_name, 'test no - zle haslo')
						EndIf

					Case $Pass[3][2] ; wartosc log_in = tomek / sprawdzenie tez dla mojego PC -  Or $Pass[1][0] = $Pass[2][3] /
						If($Pass[1][2] == $Pass[4][2] Or $Pass[1][2] == $Pass[4][4]) And ($Pass[1][0] = $Pass[2][4] Or $Pass[1][0] = $Pass[2][3]) Then MsgBox(0, $Program_name, 'test 123 - pc ok') ;  -> login ok -> haslo ok -> PC ok -> odpal program

					Case $Pass[3][3], $Pass[3][4], $Pass[3][5], $Pass[3][6] ; wartosc log_in = urzytkownik valmet / zawezona grupa /
;~ 						Local $i,
						$Pass_New = $Pass
						_ArrayTranspose($Pass_New)
						$i = _ArraySearch($Pass_New, $Pass[1][1], 0, 0, 0, 0, 1, 3)
						If $Pass[1][2] == $Pass[4][$i] Then
							If $Pass[1][0] = $Pass[2][$i] Then
;~ 								MsgBox(0, $Program_name, 'PC: ' & $Pass[1][0] & @CRLF & 'Login: ' & $Pass[1][1] & @CRLF & 'Haslo: ' & $Pass[1][2] & @CRLF & 'Index: ' & $i) ; temp
								ExitLoop
							EndIf
						EndIf
					Case Else
;~ 						MsgBox(0, $Program_name, 'Wrong, Case Else' & @CR & 'Login: ' & $Pass[1][1] & @CR & 'Haslo: ' & $Pass[1][2]) ; temp
				EndSwitch
		EndSwitch
		If $nMsg = $Button_Enter Then ; Nadpisywanie loginu i hasla przed ponowna petla
			GUICtrlSetData($Pass[1][3], $Pass[1][7])
			GUICtrlSetData($Pass[1][4], $Pass[1][8])
		EndIf
	WEnd
;~ Until $Pass[1][5] = $Pass[1][6]
	GUIDelete($Loging_Hendle)

EndFunc   ;==>Loging

Func Excel_Open()

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
			WinSetTrans($Handle_Excel, "", 255)
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

	For $l = 0 To 1
		$Do_ex_reset = MSGBox_NOYES('Czy chcesz zresetowac excela?', 5, 'Kliknij "NO" jesli nie chcesz zresetowac Excela', 2)
		GUIDelete($Handle_MSGBox_NOYES)
		If $Do_ex_reset = 2 Or $Do_ex_reset = 9 Then
			$Do_ex_show = MSGBox_NOYES('Chcesz zobaczyc dane z Excel - ' & $Txt_Excel_name & '?', 10, 'Kliknij "YES" jesli chcesz zobaczyc Excel: ' & $Txt_Excel_name, 1)
			GUIDelete($Handle_MSGBox_NOYES)
			If $Do_ex_show = 1 Then
				WinSetTrans($Handle_Excel, "", 255)
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinMove($Txt_Excel_name, '', 50, 50, 1700, 1000)
				WinActivate($Txt_Excel_name)
				Do
					Sleep(10000)
					$Do_ex_show_yet = MSGBox_NOYES('Chcesz juz wrócic do programu: ' & $Program_name & '?', 10, 'Jesli nie klikniesz "TAK" komunikat pojawi sie ponownie za 10 sec', 1, 10)
					GUIDelete($Handle_MSGBox_NOYES)
				Until $Do_ex_show_yet = 1
				If WinExists($Epl) Then WinMove($Txt_Excel_name, '', $Epl_poz[0] + 50, $Epl_poz[1] + 100, 0, 0)
				WinSetState($Txt_Excel_name, '', @SW_HIDE)
			EndIf
		Else
			Excel_Reset()
			ExitLoop
		EndIf
	Next

	WinActivate($Program_name)
EndFunc   ;==>Excel_Do_Reset

Func Excel_Reset()

	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{HOME}')
		Sleep(100)
		For $i = 1 To 3
			ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '^{z}')
			Sleep(800)
		Next
		Sleep(1000)
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
	WinActivate($Program_name)

EndFunc   ;==>Excel_Reset

Func Tworzenie_kotwicy_many()

	Global $i_many, $Liczba_Meny_how = GUICtrlRead($Meny_how)
	If $Liczba_Meny_how > 1 Then
		HotKeySet('+d', 'Exit_Procedure')
		For $i_many = 1 To $Liczba_Meny_how
			Tworzenie_kotwicy()
			If $i_many < $Liczba_Meny_how Then
				Sleep(1000)
				WinActivate($Epl)
				AutoItSetOption('MouseCoordMode', 0)
				WinWaitActive($Epl)
				MouseClick('primary', 370, 130, 1, 0)
				AutoItSetOption('MouseCoordMode', 1)
				Send("{PGDN}")
			EndIf
		Next
	Else
		Tworzenie_kotwicy()
	EndIf
	WinActivate($Program_name)

EndFunc   ;==>Tworzenie_kotwicy_many

Func Zaznaczanie($Ustaw_Zmienna = 0, $Ustaw_Aktualna = 1, $Ustaw_Wl_strony = 1)

	Local $Kategoria = 1006
	Local $Kategoria_wartosc = 'Wszystkie kategorie'
	Local $Button_Zmienna = 2006
	Local $Button_Aktualna = 2065
	Local $Button_Wl_strony = 2005
	Local $k1, $k2, $k3, $k4
	$k1 = ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "GetCurrentSelection", "")
	$k2 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Zmienna, 'IsChecked')
	$k3 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Aktualna, 'IsChecked')
	$k4 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Wl_strony, 'IsChecked')

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

	If $k1 <> $Kategoria_wartosc Then
		ControlSend($Epl_okno_kotwiczki, '', 1006, '{down}')
		Do
			ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "SelectString", $Kategoria_wartosc)
			Sleep(2000)
			$k1 = ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "GetCurrentSelection", "")
		Until $k1 = $Kategoria_wartosc
	EndIf

	If $k2 <> $Ustaw_Zmienna Or $k3 <> $Ustaw_Aktualna Or $k4 <> $Ustaw_Wl_strony Then
		ControlSend($Epl_okno_kotwiczki, '', 2006, '{space}')
		Do
			ControlCommand($Epl_okno_kotwiczki, "", $Button_Zmienna, $Status_zmienna)
			ControlCommand($Epl_okno_kotwiczki, "", $Button_Aktualna, $Status_aktualna)
			ControlCommand($Epl_okno_kotwiczki, "", $Button_Wl_strony, $Status_wl_strony)
			Sleep(2000)
			$k2 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Zmienna, 'IsChecked')
			$k3 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Aktualna, 'IsChecked')
			$k4 = ControlCommand($Epl_okno_kotwiczki, "", $Button_Wl_strony, 'IsChecked')
		Until $k2 = $Ustaw_Zmienna And $k3 = $Ustaw_Aktualna And $k4 = $Ustaw_Wl_strony
	EndIf

EndFunc   ;==>Zaznaczanie

Func MSGBox_ok($Txt_in_MSGBox_ok, $Time_up_MSGBox_ok, $Txt_Tip_MSGBox_ok = '', $Tip_icon_MSGBox_ok = 0, $9 = 9, $Czcia_MSGBox_ok = 10, $Txt_Button_MSGBox_ok = 'OK')

	Local $SEC_A, $SEC_S = @SEC, $SEC_E
	#Region ### START Koda GUI section ### Form=C:\Users\glitkaczda\Desktop\Programowanie\gui_tworzenie_kotwiczk_ok_1.2i.kxf
	Global $Okno_Tworzenie_kotwiczki_MSGBox_ok = GUICreate($Program_name & ' INFO', 339, 150)
	GUISetFont(8, 400, 0, "Showcard Gothic")
	GUISetBkColor(0x313131)
	GUICtrlCreateLabel($Txt_in_MSGBox_ok, 18, 16, 302, 58, BitOR($SS_CENTER, $SS_NOPREFIX))
	Global $Tip_MSGBox_ok = GUICtrlSetTip(-1, $Txt_Tip_MSGBox_ok, 'Info!', $Tip_icon_MSGBox_ok, 1)
	GUICtrlSetFont(-1, $Czcia_MSGBox_ok, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	Global $Button_Ok = GUICtrlCreateButton($Txt_Button_MSGBox_ok, 124, 96, 91, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	Do
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Return($9)
			Case $Button_Ok
				Return(1)
		EndSwitch
		$SEC_A = @SEC
		$SEC_E = $SEC_A - $SEC_S
	Until $SEC_E >= $Time_up_MSGBox_ok

EndFunc   ;==>MSGBox_ok

Func MSGBox_KO($Txt_in_MSGBox_KO, $Time_up_MSGBox_KO, $Txt_Tip_MSGBox_KO = '', $Tip_icon_MSGBox_KO = 0, $9 = 9, $Czcia_MSGBox_KO = 18, $Txt_Button_MSGBox_KO = 'OK')

	Local $SEC_A, $SEC_S = @SEC, $SEC_E
	#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_tworzenie_kotwiczk_ok_big1.1i.kxf
	Global $Okno_Tworzenie_kotwiczki = GUICreate("Program_name" & ' INFO', 416, 172)
	GUISetFont(8, 400, 0, "Showcard Gothic")
	GUISetBkColor(0x313131)
	GUICtrlCreateLabel($Txt_in_MSGBox_KO, 16, 16, 382, 90, BitOR($SS_CENTER, $SS_NOPREFIX))
	Global $Tip_MSGBox_KO = GUICtrlSetTip(-1, $Txt_Tip_MSGBox_KO, 'Info!', $Tip_icon_MSGBox_KO, 1)
	GUICtrlSetFont(-1, $Czcia_MSGBox_KO, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	Global $Button_Ok = GUICtrlCreateButton($Txt_Button_MSGBox_KO, 160, 120, 91, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	Do
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Return($9)
			Case $Button_Ok
				Return(1)
		EndSwitch
		$SEC_A = @SEC
		$SEC_E = $SEC_A - $SEC_S
	Until $SEC_E >= $Time_up_MSGBox_KO

EndFunc   ;==>MSGBox_KO

Func MSGBox_yesno($Txt_in_MSGBox_yesno, $Time_up_MSGBox_yesno, $Txt_Tip_MSGBox_yesno = '', $Tip_icon_MSGBox_yesno = 0, $9 = 9, $Czcia_MSGBox_yesno = 10)

	Local $SEC_A, $SEC_S = @SEC, $SEC_E
	#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_tworzenie_kotwiczk_yes_no_1.1i.kxf
	Global $Okno_Tworzenie_kotwiczki_2 = GUICreate($Program_name & ' ASK', 326, 169)
	GUISetFont(8, 400, 0, "Showcard Gothic")
	GUISetBkColor(0x313131)
	GUICtrlCreateLabel($Txt_in_MSGBox_yesno, 19, 16, 286, 82, BitOR($SS_CENTER, $SS_NOPREFIX))
	Global $Tip_MSGBox_yesno = GUICtrlSetTip(-1, $Txt_Tip_MSGBox_yesno, 'Info!', $Tip_icon_MSGBox_yesno, 1)
	GUICtrlSetFont(-1, $Czcia_MSGBox_yesno, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	Global $Button_yes = GUICtrlCreateButton("YES", 33, 120, 123, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_no = GUICtrlCreateButton("NO", 169, 120, 123, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	Do
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Return($9)
			Case $Button_yes
				Return(1)
			Case $Button_no
				Return(2)
		EndSwitch
		$SEC_A = @SEC
		$SEC_E = $SEC_A - $SEC_S
	Until $SEC_E >= $Time_up_MSGBox_yesno

EndFunc   ;==>MSGBox_yesno

Func MSGBox_NOYES($Txt_in_MSGBox_NOYES, $Time_up_MSGBox_NOYES, $Txt_Tip__MSGBox_NOYES = '', $Tip_icon_MSGBox_NOYES = 0, $9 = 9, $Czcia_MSGBox_NOYES = 18)

	Local $SEC_A, $SEC_S = @SEC, $SEC_E
	#Region ### START Koda GUI section ### Form=c:\users\glitkaczda\desktop\programowanie\gui_tworzenie_kotwiczk_yes_no_big1.1i.kxf
	Global $Okno_Tworzenie_kotwiczki_3 = GUICreate("Program_name" & ' ASK', 416, 177)
	GUISetFont(8, 400, 0, "Showcard Gothic")
	GUISetBkColor(0x313131)
	GUICtrlCreateLabel($Txt_in_MSGBox_NOYES, 24, 16, 366, 106, BitOR($SS_CENTER, $SS_NOPREFIX))
	Global $Tip_MSGBox_NOYES = GUICtrlSetTip(-1, $Txt_Tip__MSGBox_NOYES, 'Info!', $Tip_icon_MSGBox_NOYES, 1)
	GUICtrlSetFont(-1, $Czcia_MSGBox_NOYES, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0x569557)
	Global $Button_yes = GUICtrlCreateButton("YES", 66, 128, 123, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	Global $Button_no = GUICtrlCreateButton("NO", 226, 128, 123, 33)
	GUICtrlSetFont(-1, 10, 800, 0, "Showcard Gothic")
	GUICtrlSetColor(-1, 0xFFFFFF)
	GUICtrlSetBkColor(-1, 0x569557)
	GUICtrlSetCursor(-1, 0)
	GUISetState(@SW_SHOW)
	#EndRegion ### END Koda GUI section ###

	Do
		$nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				Return($9)
			Case $Button_yes
				Return(1)
			Case $Button_no
				Return(2)
		EndSwitch
		$SEC_A = @SEC
		$SEC_E = $SEC_A - $SEC_S
	Until $SEC_E >= $Time_up_MSGBox_NOYES

EndFunc   ;==>MSGBox_NOYES

Func Licence()

	Global $licence[7] = [@MDAY, StringMid($Epl_script, 20, 2), @MON, StringMid($Epl_script, 32, 2), @YEAR, StringMid($Epl_script, 36, 2),13]

	If $licence[3] < $licence[6] And Int(StringRight($licence[4], 2)) <= $licence[5] And ($licence[2] < $licence[3] Or ($licence[2] = $licence[3] And $licence[0] <= $licence[1]))  Then
		Date()
		$i = int($licence[3])
		$licence[3] = $Date[$i]
		MSGBox_KO('The license is current to' & @CRLF & $licence[1] & " " & $licence[3] & " 20" & $licence[5], 4, '', 0, 9, 20)
		GUIDelete($Handle_MSGBox_KO)
	Else
		MSGBox_KO('Your license' & @CRLF & 'is not current', 8, '', 0, 9, 26)
		GUIDelete($Handle_MSGBox_KO)
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

Func Date()
	Global $Date[13] = _
			[0, 'January', 'February', 'March', 'April', _
			'May', 'June', 'July', 'August', _
			'September', 'October', 'November', 'December']
EndFunc   ;==>Date
