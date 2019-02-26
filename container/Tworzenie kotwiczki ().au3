Tworzenie_kotwiczki ()
Func Tworzenie_kotwiczki ()

$t1 = 20000
;~ t1 - czas ³adowania okna znaku wype³niacza (8000 dla max 10 zmiennych
$t2 = 3000
;~ t2 - czas ukazywania siê wszystkich w³aœciwoœci i w³aœciwoœci strony w oknie znaku wype³niacza po ich zafajkowaniu
$tc = 2000
;~ tc - copy time


;~ 1. Skopiowanie nazw pe³nych do excela
WinActivate("EPLAN Electric P8 2.7")
   Send("^{a}")
   Send("!{t}")
   Send("{o}")
   Send("!{p}{o}{w}")
   Sleep($t1)
   Send("!{k}")
   Send("{TAB}")
   Send("^+{RIGHT}")
   Send("^+{F10}")
   Send("{p}")
   Send("!{p}{o}")
   Sleep($t2)
   Send("!{k}")
   Send("{TAB}")
   Send("^{c}")

;~ 2. Usuniêcie nazw pe³nych w excelu
WinActivate("Usuwanie_nazw_pe³nych")
   Send("^{HOME}")
   Send("^{v}")
   Send("+{Enter}")
   Send("{RIGHT}")
   Send("^+{UP}")
   Send("+{DOWN}")
   Send("^{c}")
   Sleep($tc)

;~ 3. Umieszczenie nazw wyœwietlanych i kopiowanie w³aœciwoœci
WinActivate("W³aœciwoœci (symbol graficzny)")
   Send("{RIGHT}")
   Send("{down}")
   Send("^{v}")
   Send("^{TAB}")
   Sleep($tc)
   Send("^{a}")
   Sleep($tc)
   Send("^{c}")
   Sleep($tc)

;~ 4. Zamiana W³aœciwoœci na prawid³owe w excelu
WinActivate("Usuwanie_nazw_pe³nych")
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

;~ 5. Stworzenie nowego obiektu w³aœciwoœci i przypo¿¹dkowanie prawid³owych zmiennych
WinActivate("W³aœciwoœci (symbol graficzny)")
   Send("^+{F10}")
   Send("{o}")
   Send("^{v}")
   Sleep(3000)

;~ 6. Nazwanie kotwiczki
   Send("{TAB}")
   Send("!{n}")
   Send("PREPLANNING")
   Sleep(5000)
   Send("{Enter}")
EndFunc