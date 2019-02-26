#include <ButtonConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#Region ### START Koda GUI section ### Form=
Global $Okno_Tworzenie_kotwiczki = GUICreate("Tworzenie kotwiczki z Darkiem :D", 381, 192, 192, 124)
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
Global $Button_W_Wac = GUICtrlCreateButton("W Waclawa", 23, 95, 211, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
Global $Button_Jeszcze_wiecej = GUICtrlCreateButton("Jeszcze wiecej", 22, 137, 211, 33)
GUICtrlSetFont(-1, 10, 800, 0, "MS Sans Serif")
GUICtrlSetBkColor(-1, 0xA0A0A0)
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###


Open_Kotwiczka()
WinActivate("Tworzenie kotwiczki z Darkiem :D")
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			WinActivate("Kotwiczka")
			Czyszczenie_excela()
			WinActivate("Kotwiczka")
			Send("!{F4}")
			Sleep(100)
			Send("{RIGHT}")
			Send("{Enter}")
			MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Wszystko pozamykane, dziekuje ;*", 5)
			Exit

		Case $Button_1
			$Buton_x = 1
			Tworzenie_kotwicy()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")
			MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Kotwiczka stworzona (mam nadzieje) ;*", 10)

		Case $Button_2
			$Buton_x = 2
			Tworzenie_kotwicy()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")
			MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Kotwiczka stworzona (mam nadzieje) ;*", 10)

		Case $Button_3
			$Buton_x = 3
			Tworzenie_kotwicy()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")
			MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Kotwiczka stworzona (mam nadzieje) ;*", 10)

		Case $Button_4
			$Buton_x = 4
			Tworzenie_kotwicy()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")
			MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Kotwiczka stworzona (mam nadzieje) ;*", 10)

		Case $Button_W_Wac
			$Buton_x = 5
			Tworzenie_kotwicy()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")
			MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Kotwiczka stworzona (mam nadzieje) ;*", 10)

		Case $Button_Jeszcze_wiecej
			$Buton_x = 8
			Tworzenie_kotwicy()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")
			MsgBox(0,"Tworzenie kotwiczki z Darkiem :D","Kotwiczka stworzona (mam nadzieje) ;*",10)


		Case $Button_Excel_Open
			Open_Kotwiczka()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")

		Case $Button_Excel_Reset
			Czyszczenie_excela()
			WinActivate("Tworzenie kotwiczki z Darkiem :D")

		Case $Button_Excel_Close
			Czyszczenie_excela()
			WinActivate("Kotwiczka")
			Send("!{F4}")
			Sleep(100)
			Send("{RIGHT}")
			Send("{Enter}")

	EndSwitch
WEnd


Func Open_Kotwiczka()
	MsgBox(0, "Tworzenie kotwiczki z Darkiem :D", "Poczekaj, az wyskoczy kolejne okienko. Ok?")
	Send("#r")
	Sleep(250)
	Send("Excel")
	Send("{Enter}")
	WinWaitActive("Excel")
	Send("!{o}")
	Send("!{o 2}")
	Send("C:\Users\glitkaczda\Desktop\Program Darka do tworzenia kotwiczek\Kotwiczka.xlsx")
	Send("{Enter}")
	Sleep(5000)
	MsgBox(0, "Excel info", "Excel jest gotowy do dzialania", 7)
EndFunc   ;==>Open_Kotwiczka

Func Close_Kotwiczka ()
WinActivate("Kotwiczka")
			Send("!{F4}")
			Sleep(100)
			Send("{RIGHT}")
			Send("{Enter}")
EndFunc

Func Czyszczenie_excela()
	WinActivate("Kotwiczka")
	Send("^{HOME}")
	Send("^{z 20}")
	Send("^{HOME}")
	Sleep(100)
	MsgBox(0, "Excel info", "Excel czysty", 0.5)
EndFunc   ;==>Czyszczenie_excela

Func Tworzenie_kotwicy()

	$pis = $Buton_x ;~ Ilosc kart na obwodówce

	$t1 = (12000 * ((1.0 * $pis))) ;~ t1 - czas ladowania okna znaku wypelniacza dla jednego sygnalu

	$t2 = (7000 * ((0.3 * $pis) + 1)) ;~ t2 - czas ukazywania sie wszystkich wlasciwosci i wlasciwosci strony w oknie znaku wypelniacza po ich zafajkowaniu

	$t3 = (3000 * ((0.3 * $pis) + 1)) ;~ t3 - czas przenoszenia zmiennych

	$t4 = (15000 * ((0.6 * $pis) + 1)) ;~ t4 - czas ukazywania sie wszystkich wlasciwosci i wlasciwosci strony w oknie znaku wypelniacza po ich zafajkowaniu

	$tc = 2000 ;~ tc - copy time

	$ta = 1000 ;~ ta - approve time


;~ 1. Skopiowanie nazw pelnych do excela
	WinActivate("EPLAN Electric P8 2.7")
	Send("^{a}")
	Send("!{t}")
	Send("{o}")
	Sleep($t1)
	Send("!{p}{o}{w}")
	Sleep($t2)
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
	Send("^{HOME}")
	Send("^{v}")
	Send("+{Enter}")
	Send("{RIGHT}")
	Send("^+{UP}")
	Send("+{DOWN}")
	Send("^{c}")
	Sleep($tc)

;~ 3. Umieszczenie nazw wyswietlanych i kopiowanie wlasciwosci
	WinActivate("W³aœciwoœci (symbol graficzny)")
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
	WinActivate("W³aœciwoœci (symbol graficzny)")
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
EndFunc   ;==>Tworzenie_kotwicy
