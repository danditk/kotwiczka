#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <AutoItConstants.au3>
#include <MsgBoxConstants.au3>

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
Global $Button_2 = GUICtrlCreateButton("2", 80, 56, 43, 33)
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
Global $Button_Jeszcze_wiecej = GUICtrlCreateButton("Jeszcze wiecej", 23, 95, 211, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_proces_stop = GUICtrlCreateButton("Hope quick exit", 22, 137, 211, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $Ask_ex_open, $Txt_ex_close, $login

$Ask_ex_open = 'Czy chcesz otworzyc Excel - "Kotwiczka"?'
$Txt_ex_close = 'Excel - "Kotwiczka" nie jest otwarty'

Login_user()
Excel_Open()
Activate_program_name()


While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			Excel_Close()
			Bye()
			Exit

		Case $Button_1
			$Button_x = 0.8
			Tworzenie_kotwicy()

		Case $Button_2
			$Button_x = 2
			Tworzenie_kotwicy()

		Case $Button_3
			$Button_x = 3
			Tworzenie_kotwicy()

		Case $Button_4
			$Button_x = 4
			Tworzenie_kotwicy()

		Case $Button_proces_stop
			Process_stop()

		Case $Button_Jeszcze_wiecej
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

	Global $Sciezka_ex, $haslo_moje, $haslo_tomek
	Local $Sciezka_cz1, $Sciezka_cz2
	$haslo_moje = "no"
	$haslo_tomek = "123"
	$login = InputBox("Tworzenie kotwiczki z Darkiem :D","Prosze, wpisz swój login","user_login")
	$Sciezka_cz1 = 'C:\Users\'
	$Sciezka_cz2 = '\Desktop\Program Darka do tworzenia kotwiczek\Kotwiczka.xlsx'

	If $login = $haslo_moje Then
		$login = "glitkaczda"
	ElseIf $login = $haslo_tomek Then
		$login = "glinoconto"
	EndIf
	$Sciezka_ex = $Sciezka_cz1 & $login & $Sciezka_cz2

EndFunc

Func Activate_program_name()

	WinActivate($Program_name)

EndFunc

Func Activate_program_name_err()

	MsgBox(0,$Program_name,"Chyba cos poszlo nie tak",3)
	Activate_program_name()

EndFunc

Func Bye()

	MsgBox(0, $Program_name, "Wszystko pozamykane, dziekuje " & @CRLF & @CRLF & "BYE!", 3)

EndFunc   ;==>Bye

Func Excel_Open()

	Local $ok = MsgBox(1, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?")
	If $ok = 1 Then
		Send("#r")
		Sleep(250)
		Send("Excel")
		Send("{Enter}")
		WinWaitActive("Excel")
		Send("!{o}")
		Send("!{o 2}")
		Send($Sciezka_ex)
		Send("{Enter}")
		If WinWaitActive("Kotwiczka","",15) Then
			MsgBox(0, "Excel info", "Excel jest gotowy do dzialania", 7)
			Activate_program_name()
		Else
			Activate_program_name_err()
		EndIf
	EndIf

EndFunc   ;==>Excel_Open

Func Excel_Close()

	If WinExists("Kotwiczka") Then
		WinActivate("Kotwiczka")
		WinWaitActive("Kotwiczka","",20)
		WinClose("Kotwiczka")
		Sleep(500)
		Send("{n}")
		MsgBox(0, "Excel", "Excel zostal zamkniety", 2)
		Activate_program_name()
	EndIf

EndFunc   ;==>Excel_Close

Func Excel_Close_Button()

	If WinExists("Kotwiczka") Then
		WinActivate("Kotwiczka")
		WinClose("Kotwiczka")
		Sleep(500)
		Send("{n}")
		MsgBox(0, "Excel", "Excel zostal zamkniety", 3)
		Activate_program_name()
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
	EndIf

EndFunc   ;==>Excel_Close_Button

Func Excel_Reset()

	If WinExists("Kotwiczka") Then
		WinActivate("Kotwiczka")
		Send("^{HOME}")
		Send("^{z 20}")
		Send("^{HOME}")
		Sleep(100)
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

	If WinExists("Kotwiczka") Then

		Local $Epl = "EPLAN Electric P8 2.7"
		If WinExists($Epl) Then

			MsgBox(0, $Program_name, "Poczekaj, az wyskoczy kolejne okienko. Ok?" & @CRLF & "Jak nic sie nie bedzie dzialo to znaczy, ze to TY cos zrobiles nie tak.")
			$pis = $Button_x ;~ Ilosc kart na obwodówce
			$t2 = (7000 * ((0.3 * $pis) + 1)) ;~ t2 - czas ukazywania sie wszystkich wlasciwosci i wlasciwosci strony w oknie znaku wypelniacza po ich zafajkowaniu
			$t3 = (3000 * ((0.3 * $pis) + 1)) ;~ t3 - czas przenoszenia zmiennych
			$t4 = (15000 * ((0.6 * $pis) + 1)) ;~ t4 - czas ukazywania sie wszystkich wlasciwosci i wlasciwosci strony w oknie znaku wypelniacza po ich zafajkowaniu
			$tc = 2000 ;~ tc - copy time
			$ta = 1000 ;~ ta - approve time

	;~ 1. Skopiowanie nazw pelnych do excela
			Local $Epl_okno_kotwiczki = "W³aœciwoœci (symbol graficzny)"
			WinActivate($Epl)
			Send("^{a}")
			Send("!{t}")
			Send("{o}")
			WinWaitActive($Epl_okno_kotwiczki,"",60)
			If WinExists($Epl_okno_kotwiczki) Then
				Sleep($ta)
				Send("!{p}{o}{w}")
				Sleep($t4)
				Send("!{k}")
				Send("{w}")
				Send("!{k}")
				Send("{TAB}")
				Send("^+{RIGHT}")
				Send("^+{F10}")
				Send("{p}")
				Sleep($t3)
				Send("!{p}{o}")
				Sleep($t4)
				Send("!{k}")
				Send("{TAB}")
				Send("^{c}")


		;~ 2. Usuniecie nazw pelnych w excelu
				WinActivate("Kotwiczka")
				WinWaitActive("Kotwiczka")
				Send("^{HOME}")
				Send("^{v}")
				Send("+{Enter}")
				Send("{RIGHT}")
				Send("^+{UP}")
				Send("+{DOWN}")
				Send("^{c}")
				Sleep($tc)

		;~ 3. Umieszczenie nazw wyswietlanych i kopiowanie wlasciwosci
				WinActivate($Epl_okno_kotwiczki)
				WinWaitActive($Epl_okno_kotwiczki)
				Send("{RIGHT}")
				Send("{DOWN}")
				Send("^{v}")
				Sleep($tc)
				Send("^{TAB}")
				Sleep(500)
				Send("^{a}")
				Send("^{c}")
				Sleep($tc)

		;~ 4. Zamiana Wlasciwosci na prawidlowe w excelu
				WinActivate("Kotwiczka")
				WinWaitActive("Kotwiczka")
				Send("{F5}")
				Send("{z}")
				Send("{Enter}")
				Send("^{v}")
				Send("+{Enter}")
				Send("{RIGHT}")
				Send("^+{UP}")
				Send("+{DOWN}")
				Send("^{c}")
				Sleep($tc)

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

				$Do_ex_reset = MsgBox(4,$Program_name,"Czy chcesz zresetowac excela?",5)
				If $Do_ex_reset <> 7 Then Excel_Reset()
				Activate_program_name()
			Else
				Activate_program_name_err()
			EndIf
		Else
			MsgBox(1,$Program_name,"Program " & $Epl & " nie zostal wlaczony." & @CRLF & "Wlacz " & $Epl & "kliknij na strone i sprobuj ponownie")
		EndIf
	Else
		MsgBox(0, $Program_name, $Txt_ex_close)
		$Do_ex_open = MsgBox(1, $Program_name, $Ask_ex_open)
		If $Do_ex_open = 1 Then
			Excel_Open()
		Else
			MsgBox(0, $Program_name, "To moj pierwszy program." & @CRLF & "Nie posiadam jeszcze dostatecznej wiedzy, by zrobic to bez Excela")
		EndIf
	EndIf

EndFunc   ;==>Tworzenie_kotwicy

Func Process_stop()
	Send("{BREAK}")
EndFunc
