Tworzenie_kotwiczki ()
Func Tworzenie_kotwiczki ()

$t1 = 20000
;~ t1 - czas ładowania okna znaku wypełniacza (8000 dla max 10 zmiennych
$t2 = 3000
;~ t2 - czas ukazywania się wszystkich właściwości i właściwości strony w oknie znaku wypełniacza po ich zafajkowaniu
$tc = 2000
;~ tc - copy time


;~ 1. Skopiowanie nazw pełnych do excela
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

;~ 2. Usunięcie nazw pełnych w excelu
WinActivate("Usuwanie_nazw_pełnych")
   Send("^{HOME}")
   Send("^{v}")
   Send("+{Enter}")
   Send("{RIGHT}")
   Send("^+{UP}")
   Send("+{DOWN}")
   Send("^{c}")
   Sleep($tc)

;~ 3. Umieszczenie nazw wyświetlanych i kopiowanie właściwości
WinActivate("Właściwości (symbol graficzny)")
   Send("{RIGHT}")
   Send("{down}")
   Send("^{v}")
   Send("^{TAB}")
   Sleep($tc)
   Send("^{a}")
   Sleep($tc)
   Send("^{c}")
   Sleep($tc)

;~ 4. Zamiana Właściwości na prawidłowe w excelu
WinActivate("Usuwanie_nazw_pełnych")
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

;~ 5. Stworzenie nowego obiektu właściwości i przypożądkowanie prawidłowych zmiennych
WinActivate("Właściwości (symbol graficzny)")
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