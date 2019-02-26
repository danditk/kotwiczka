#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>
#include <Excel.au3>

#Region ### START Koda GUI section ### Form=
Global $Program_name = "Tworzenie kotwiczki z Darkiem :D"
Global $Okno_Tworzenie_kotwiczki = GUICreate($Program_name, 381, 192, 192, 124)
Global $Tekst_Ilosc_kart = GUICtrlCreateLabel("Ilosc kart na obwodowce", 24, 16, 214, 26, $SS_CENTER)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $Text_excel = GUICtrlCreateLabel("Excel", 264, 16, 94, 26, $SS_CENTER)
GUICtrlSetFont(-1, 12, 800, 0, "MS Sans Serif")
Global $Button_1 = GUICtrlCreateButton("1", 24, 56, 43, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_2 = GUICtrlCreateButton("2", 80, 56, 43, 33);)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_4 = GUICtrlCreateButton("4", 192, 56, 43, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_3 = GUICtrlCreateButton("3", 136, 56, 43, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_Excel_Open = GUICtrlCreateButton("Open", 264, 56, 91, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_Excel_Reset = GUICtrlCreateButton("Reset", 263, 96, 91, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_Excel_Close = GUICtrlCreateButton("Close", 264, 136, 91, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_wiele_danych = GUICtrlCreateButton("Dla wielu danych", 23, 95, 211, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_proces_stop = GUICtrlCreateButton("Error - quick exit", 22, 137, 211, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $Txt_Excel_name, $Excel_full_name, $Sciezka_excela, $Program_Excel_open, $Plik_Excel_open
Global $Epl_okno_kotwiczki, $Ask_ex_open, $Txt_ex_close, $login
$Txt_Excel_name = 'Kotwiczka'
$Ask_ex_open = 'Czy chcesz otworzyc Excel - ' & $Txt_Excel_name & '?'
$Epl_okno_kotwiczki = "W³aœciwoœci (symbol graficzny)"
$Txt_ex_close = 'Excel - ' & $Txt_Excel_name & ' nie jest otwarty.'
HotKeySet('e', 'HotKey_Exit')

;~ Login_user()
Excel_Open()
Activate_program_name()

While 1

	HotKeySet('e', 'HotKey_Exit')
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Excel_Close()
			Bye()
			Exit

		Case $Button_1
			$Button_x = 1
			Tworzenie_kotwicy()

		Case $Button_proces_stop
			HotKey_Exit()

		Case $Button_wiele_danych
			$Button_x = 6
			Tworzenie_kotwicy()

		Case $Button_Excel_Open
			Excel_Open()

		Case $Button_Excel_Reset
			Excel_Reset()

		Case $Button_Excel_Close
			Excel_Close_Button()

	EndSwitch

WEnd

Func Login_user()

	HotKeySet('e', 'HotKey_Exit')
	Global $Sciezka_ex, $haslo_moje, $haslo_tomek1
	Local $Sciezka_cz1, $Sciezka_cz2, $Txt_login, $login_wrong, $login_restart
	$haslo_moje1 = "no"
	$haslo_moje2 = "danditkaczuk"
	$haslo_tomek1 = "123"
	$haslo_tomek2 = "glinoconto"
	$haslo_pawel1 = "glikozowpa"
	$haslo_lukasz1 = "gliwoliclu"
	$haslo_Ola = "ola"
	$Txt_login = "gli pc_user_login"
	$login = InputBox("Tworzenie kotwiczki z Darkiem :D", "Prosze, wpisz swój 10 literowy login", $Txt_login)
;~ 	$Sciezka_cz1 = 'C:\Users\'
;~ 	$Sciezka_cz2 = '\Desktop\Program Darka do tworzenia kotwiczek\Kotwiczka.xlsx'

	If $login = $haslo_moje1 Or $login = $haslo_moje2 Then
		$login = "glitkaczda"
	ElseIf $login = $haslo_tomek1 Or $login = $haslo_tomek2 Then
		$login = "glinoconto"
		MsgBox(0, $Program_name, "Czesc Tomasz ;D ", 1)
	ElseIf $login = $haslo_pawel1 Then
		$login = "glikozowpa"
		MsgBox(0, $Program_name, "Czesc Pewel !", 1)
	ElseIf $login = $haslo_lukasz1 Then
		$login = "gliwoliclu"
		MsgBox(0, $Program_name, "Czesc Lukasz !", 1)
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
;~ 	$Sciezka_ex = $Sciezka_cz1 & $login & $Sciezka_cz2

EndFunc   ;==>Login_user

Func Activate_program_name()

	WinActivate($Program_name)

EndFunc   ;==>Activate_program_name

Func Activate_program_name_err()

	MsgBox(0, $Program_name, "Chyba cos poszlo nie tak", 3)
	Activate_program_name()

EndFunc   ;==>Activate_program_name_err

Func Bye()

	MsgBox(0, $Program_name, "Wszystko pozamykane, dziekuje " & @CRLF & @CRLF & "BYE!", 3)

EndFunc   ;==>Bye

Func Excel_Open()

	HotKeySet('e', 'HotKey_Exit')
	Local $ok = MsgBox(1, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?")
	If $ok = 1 Then
		WinActivate($Txt_Excel_name)
		Local $ukosnik, $rozszerzenie
		$ukosnik = '\'
		$rozszerzenie = '.xlsx'
		$Excel_full_name = $Txt_Excel_name & $rozszerzenie
		$ukosnik &= $Excel_full_name
		$Sciezka_excela = @ScriptDir & $ukosnik
		$Program_Excel_open = _Excel_Open()
		$Plik_Excel_open = _Excel_BookOpen($Program_Excel_open, $Sciezka_excela)
		If WinWaitActive($Txt_Excel_name, "", 15) Then
			MsgBox(0, "Excel info", "Excel jest gotowy do dzialania", 5)
			Activate_program_name()
		Else
			Activate_program_name_err()
		EndIf
	EndIf

EndFunc   ;==>Excel_Open

Func Excel_Close()

	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		WinWaitActive($Txt_Excel_name, "", 20)
		WinClose($Txt_Excel_name)
		Sleep(500)
		Send("{n}")
		MsgBox(0, "Excel", "Excel zostal zamkniety", 2)
		Activate_program_name()
	EndIf

EndFunc   ;==>Excel_Close

Func Excel_Close_Button()

	HotKeySet('e', 'HotKey_Exit')
	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		WinClose($Txt_Excel_name)
		Sleep(500)
		Send("{n}")
		MsgBox(0, "Excel", "Excel zostal zamkniety", 3)
		Activate_program_name()
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
	EndIf

EndFunc   ;==>Excel_Close_Button

Func Excel_Reset()

	HotKeySet('e', 'HotKey_Exit')
	If WinExists($Txt_Excel_name) Then
		WinActivate($Txt_Excel_name)
		Send("^{HOME}")
		Send("^{z 50}")
		Send("^{HOME}")
		Sleep(200)
		MsgBox(0, "Excel info", "Excel czysty", 0.5)
		Activate_program_name()
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
		$Do_ex_open = MsgBox(1, $Program_name, $Ask_ex_open)
		If $Do_ex_open = 1 Then
			Excel_Open()
		EndIf

	EndIf

EndFunc   ;==>Excel_Reset

Func Tworzenie_kotwicy()

	HotKeySet('e', 'HotKey_Exit')
	If WinExists($Txt_Excel_name) Then
		Local $widocznosc, $Epl = "EPLAN Electric P8 2.7"
		If WinExists($Epl) Then
			MsgBox(0, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?" & @CRLF & "Jak nic sie nie bedzie dzialo to znaczy, ze to TY cos zrobiles nie tak.")
			$pis = $Button_x ;~ Ilosc kart na obwodówce
			$t3 = (3000 * ((0.3 * $pis) + 1)) ;~ t3 - czas przenoszenia zmiennych
			$tc = (2000*  ((1 * $pis)      )) ;~ tc - copy time
			$ta = 500 ;~ ta - approve timen

;~ 1. Skopiowanie nazw pelnych do excela
			WinActivate($Epl)
			AutoItSetOption('MouseCoordMode',0)
			Sleep(250)
			WinWaitActive($Epl)
			MouseClick('primary',370,130,1,0)
			Sleep(250)
			ControlSend($Epl,'',"[CLASSNN:AfxFrameOrView140u1]",'^{a}')
			Sleep(100)
			ControlSend($Epl,'','Afx:0000000140000000:8:0000000000010007:0000000000000010:000000000000000015','!{t}')
			ControlSend($Epl,'','Afx:0000000140000000:8:0000000000010007:0000000000000010:000000000000000015','{o}')
			WinActivate($Epl)
			If $nMsg = $Button_wiele_danych Then MsgBox(0,$Program_name,'Kliknij po ukazaniu sie okienka kotwiczki, by przyspieszyc 2 minutowe odliczanie.',120)
			WinWaitActive($Epl_okno_kotwiczki, "", 60)
			If WinActive($Epl_okno_kotwiczki) Then
				Sleep(1000)
				WinActivate($Epl_okno_kotwiczki)
				Sleep(1000)
				WinWait($Epl_okno_kotwiczki)
				Sleep(1000)
				WinWaitActive($Epl_okno_kotwiczki)
				Sleep(1000)
				WinActivate($Epl_okno_kotwiczki)
				AutoItSetOption('MouseCoordMode',0)
				MouseClick('primary',100,70,1,0)
				Sleep(100)
				MouseClick('primary',115,190,1,0)
				Sleep(100)
				MouseClick('primary',115,190,1,0)
				Sleep(100)
				MouseClick('primary',90,290,1,0)
				Sleep(100)
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
				Sleep($t3)
				$Ustaw_Zmienna = 1
				$Ustaw_Aktualna = 0
				$Ustaw_Wl_strony = 1
				Zaznaczanie()
				Send("!{k}")
				Send("{TAB}")
				Send("^{c}")
				Sleep($tc)

;~ 2. Usuniecie nazw pelnych w excelu

				ControlSend($Txt_Excel_name,'','NetUIHWND2','^{HOME}')
				Sleep($tc)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','^{v}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','+{Enter}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','{RIGHT}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','^+{UP}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','+{DOWN}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','^{c}')
				Sleep($tc)

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

				ControlSend($Txt_Excel_name,'','NetUIHWND2','{F5}')
				Sleep($ta)
				WinActivate($Txt_Excel_name)
				Sleep($ta)
				ControlSend('Go To','','EDTBX1','{z}')
				Sleep($ta)
				ControlSend('Go To','','EDTBX1','{Enter}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','{RIGHT}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','{LEFT}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','^{v}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','+{Enter}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','{RIGHT}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','^+{UP}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','+{DOWN}')
				Sleep($ta)
				ControlSend($Txt_Excel_name,'','NetUIHWND2','^{c}')
				Sleep($tc)

;~ 				ControlSend($Txt_Excel_name,'','NetUIHWND2','{F5}')
;~ 				Sleep(1000)
;~ 				WinActivate($Txt_Excel_name)
;~ 				ControlSend('Go To','','EDTBX1','{z}')
;~ 				Sleep(1000)
;~ 				ControlSend('Go To','','EDTBX1','{Enter}')
;~ 				Sleep(1000)
;~ 				ControlSend('Go To','','NetUIHWND2','Blah')


;~ 				WinActivate($Txt_Excel_name)
;~ 				WinWaitActive($Txt_Excel_name)
;~ 				Send("{F5}")
;~ 				Send("{z}")
;~ 				Send("{Enter}")
;~ 				Send("^{v}")
;~ 				Send("+{Enter}")
;~ 				Send("{RIGHT}")
;~ 				Send("^+{UP}")
;~ 				Send("+{DOWN}")
;~ 				Send("^{c}")
;~ 				Sleep($tc)

;~ 5. Stworzenie nowego obiektu wlasciwosci i przypo¿¹dkowanie prawidlowych zmiennych
				WinActivate($Epl_okno_kotwiczki)
				WinWaitActive($Epl_okno_kotwiczki)
				Send("^+{F10}")
				Send("{o}")
				Send("^{v}")
				Sleep($tc)

;~ 6. Nazwanie kotwiczki
				Send("{TAB}")
				Send("!{n}")
				Send("PREPlANNING")
				Sleep($ta)
				Send("{Enter}")

				MsgBox(0, $Program_name, "Kotwiczka stworzona ( mam nadzieje ;D )", 10)

				$Do_ex_reset = MsgBox(4, $Program_name, "Czy chcesz zresetowac excela?", 5)
				If $Do_ex_reset <> 7 Then Excel_Reset()
				Activate_program_name()
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
	HotKeySet('e', 'HotKey_Exit')
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
		HotKeySet('e', 'HotKey_Exit')
		ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "SelectString", $Kategoria_wartosc)
		Sleep(500)
		Local $k1
		$k1 = ControlCommand($Epl_okno_kotwiczki, "", $Kategoria, "GetCurrentSelection", "")
	Until $k1 = $Kategoria_wartosc

	Do
		HotKeySet('e', 'HotKey_Exit')
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
EndFunc   ;==>Process_stop