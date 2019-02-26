#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>

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


Global $Txt_Excel_name, $Excel_full_name, $Sciezka_excela, $Program_Excel_open, $Plik_Excel_open
Global $Epl, $Epl_poz, $Epl_okno_kotwiczki_poz, $Epl_okno_kotwiczki, $Ask_ex_open, $Txt_ex_close, $login
Global $Liczba_Meny_how
$Epl = "EPLAN Electric P8 2.7"
$Epl_poz = WinGetPos($Epl)
$Txt_Excel_name = 'Kotwiczka'
$Ask_ex_open = 'Czy chcesz otworzyc Excel - ' & $Txt_Excel_name & '?'
$Epl_okno_kotwiczki = "W³aœciwoœci (symbol graficzny)"
$Txt_ex_close = 'Excel - ' & $Txt_Excel_name & ' nie jest otwarty.'
HotKeySet('+!e', 'HotKey_Exit')

;~ Login_user()
Excel_Open()
WinActivate($Program_name)

While 1

	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Excel_Close()
			MsgBox(0, $Program_name, "Wszystko pozamykane, dziekuje " & @CRLF & @CRLF & "BYE!", 3)
			Exit

		Case $Button_Start
			Tworzenie_kotwicy()

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

Func Login_user()

	Global $Sciezka_ex, $haslo_moje[1], $haslo_tomek[1]
	Local $Sciezka_cz1, $Sciezka_cz2, $Txt_login, $login_wrong, $login_restart
	$haslo_moje[0] = "no"
	$haslo_moje[1] = "danditkaczuk"
	$haslo_tomek[0] = "123"
	$haslo_tomek[1] = "glinoconto"
	$haslo_pawel1 = "glikozowpa"
	$haslo_lukasz1 = "gliwoliclu"
	$haslo_Ola = "ola"
	$Txt_login = "gli pc_user_login"
	$login = InputBox("Tworzenie kotwiczki z Darkiem :D", "Prosze, wpisz swój 10 literowy login", $Txt_login)
	If $login = $haslo_moje[0] Or $login = $haslo_moje[1] Then
		$login = "glitkaczda"
	ElseIf $login = $haslo_tomek[0] Or $login = $haslo_tomek[1] Then
		$login = "glinoconto"
		MsgBox(0, $Program_name, "Czesc Tomasz ;D ", 2)
	ElseIf $login = $haslo_pawel1 Then
		$login = "glikozowpa"
		MsgBox(0, $Program_name, "Czesc Pewel !", 2)
	ElseIf $login = $haslo_lukasz1 Then
		$login = "gliwoliclu"
		MsgBox(0, $Program_name, "Czesc Lukasz !", 2)
	ElseIf $login = $haslo_Ola Then
		$login = "glitkaczda"
		MsgBox(0, $Program_name, "Czesc kochanie ;*", 5)
	ElseIf $login = $Txt_login Then
		MsgBox(0, $Program_name, "Nie wpisales hasla", 2)
		$login_wrong = 1
	ElseIf StringLen($login) <> 10 Then
		MsgBox(0, $Program_name, "Wpisales zly login lub chcesz wyjsc")
		$login_wrong = 1
	EndIf
	If $login_wrong = 1 Then
		$login_restart = MsgBox(4, $Program_name, "Chcesz sprobowac jeszcze raz?")
		If $login_restart = 6 Then
			Login_user()
		Else
			Exit
		EndIf
	EndIf

EndFunc   ;==>Login_user

Func Activate_program_name_err()

	MsgBox(0, $Program_name, "Chyba cos poszlo nie tak", 3)
	WinActivate($Program_name)

EndFunc   ;==>Activate_program_name_err

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
					MsgBox(0, $Program_name, "Program " & $Epl & " nie zostal wlaczony." & @CRLF & "Pamietaj by wlaczyc " & $Epl & " i kliknac na strone.", 3)
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
		WinSetState($Txt_Excel_name, '', @SW_SHOW)
		WinMove($Txt_Excel_name, '', $Epl_poz[0] + 200, $Epl_poz[1] + 100, 1700, 800)
		WinActivate($Txt_Excel_name)
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
		MsgBox(0, "Excel info", "Excel czysty", 1)
		WinActivate($Program_name)
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
		$Do_ex_open = MsgBox(1, $Program_name, $Ask_ex_open)
		If $Do_ex_open = 1 Then
			Excel_Open()
		EndIf

	EndIf

EndFunc   ;==>Excel_Reset

Func Tworzenie_kotwicy()

	Local $Do_ex_reset, $Handle_Epl_okno_kotwiczki, $Handle_Excel, $Przez_Excel, $Przez_Epl_okno_kotwiczki
	If WinExists($Txt_Excel_name) Then
		If WinExists($Epl) Then
			MsgBox(0, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?" & @CRLF & "Jak nic sie nie bedzie dzialo to znaczy, ze to TY cos zrobiles nie tak.", 5)
			$tc = 2000 ;~ tc - copy time
			$ta = 250 ;~ ta - approve timen
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
			WinWait($Epl_okno_kotwiczki, '', 5)
			$Handle_Excel = WinGetHandle($Txt_Excel_name)
			WinSetTrans($Handle_Excel, "", 0)
			WinSetState($Txt_Excel_name, '', @SW_SHOW)
			WinActivate($Txt_Excel_name)
			For $i = 0 to 255
				WinSetTrans($Handle_Excel, "", $i)
				Sleep(10)
			Next
			WinSetTrans($Handle_Excel, "", 255)
			WinSetOnTop($Epl_okno_kotwiczki, '', 1)
;~ 			If $nMsg = $Button_wiele_danych Then MsgBox(0, $Program_name, 'Kliknij po ukazaniu sie okienka kotwiczki, by przyspieszyc 2 minutowe odliczanie.', 120)
			WinActivate($Epl_okno_kotwiczki)
			WinWaitActive($Epl_okno_kotwiczki, "", 60)
			If WinActive($Epl_okno_kotwiczki) Then
				Sleep(1000)
				$Epl_okno_kotwiczki_poz = WinGetPos($Epl_okno_kotwiczki)
				WinActivate($Txt_Excel_name)
				WinMove($Txt_Excel_name, '', $Epl_okno_kotwiczki_poz[0] + 100, $Epl_okno_kotwiczki_poz[1] + 100, $Epl_okno_kotwiczki_poz[2], $Epl_okno_kotwiczki_poz[3], 10)
				For $i = 255 to $Przez_Excel Step -1
					WinSetTrans($Handle_Excel, "", $i)
					Sleep(10)
				Next
				WinSetTrans($Handle_Excel, "", $Przez_Excel)
				WinActivate($Epl_okno_kotwiczki)
				$Handle_Epl_okno_kotwiczki = WinGetHandle($Epl_okno_kotwiczki)
				For $i = 255 to $Przez_Epl_okno_kotwiczki Step -1
					WinSetTrans($Handle_Epl_okno_kotwiczki, "", $i)
					Sleep(20)
				Next
				WinSetTrans($Handle_Epl_okno_kotwiczki, "", $Przez_Epl_okno_kotwiczki)
				WinSetState($Txt_Excel_name, '', @SW_HIDE)
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
				Global $Ustaw_Zmienna = 0
				Global $Ustaw_Aktualna = 1
				Global $Ustaw_Wl_strony = 1
				Zaznaczanie()
				Sleep(100)
				WinActivate($Epl_okno_kotwiczki)
				Send("!{k}")
				Send("{TAB}")
				Send("^{a}")
				Send("^+{F10}")
				Send("{p}")
				$Ustaw_Zmienna = 1
				$Ustaw_Aktualna = 0
				$Ustaw_Wl_strony = 1
				Zaznaczanie()
				Send("!{k}")
				Send("{TAB}")
				Send("^{c}")
				Sleep($tc)

;~ 2. Usuniecie nazw pelnych w excelu
				WinSetState($Txt_Excel_name, '', @SW_SHOW)
				WinActivate($Txt_Excel_name)
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
				ControlSend($Txt_Excel_name, '', 'NetUIHWND2', '{F5}')
				Sleep(300)
				WinSetTrans('Go To', "", $Przez_Excel)
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
				Sleep($ta)
				Send("{Enter}")
				WinMove($Txt_Excel_name, '', $Epl_poz[0] + 50, $Epl_poz[1] + 100, 0, 0)
				WinSetTrans($Handle_Excel, "", 255)
				WinSetOnTop($Epl_okno_kotwiczki, '', 0)
				MsgBox(0, $Program_name, "Kotwiczka stworzona ( mam nadzieje ;D )", 3)
				Excel_Do_Reset()
				If $licznik <> $licznik_x Then Excel_Reset()
				WinActivate($Program_name)
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

EndFunc   ;==>Tworzenie_kotwicy

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

Func HotKey_Exit()
	Exit
EndFunc   ;==>HotKey_Exit