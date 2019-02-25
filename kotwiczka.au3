#include <StaticConstants.au3>
#include <WindowsConstants.au3>

$pis = $Buton_x ;~ Ilosc kart na obwodówce

$t1 = (12000 * ((1.0* $pis)  ));~ t1 - czas ladowania okna znaku wypelniacza dla jednego sygnalu

$t2 = (7000  * ((0.3* $pis)+1));~ t2 - czas ukazywania sie wszystkich wlasciwosci i wlasciwosci strony w oknie znaku wypelniacza po ich zafajkowaniu

$t3 = (3000  * ((0.3* $pis)+1));~ t3 - czas przenoszenia zmiennych

$t4 = (15000 * ((0.6* $pis)+1));~ t4 - czas ukazywania sie wszystkich wlasciwosci i wlasciwosci strony w oknie znaku wypelniacza po ich zafajkowaniu

$tc =  2000 ;~ tc - copy time

$ta =  1000 ;~ ta - approve time


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