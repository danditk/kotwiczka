Tworzenie_kotwiczki ()
Func Tworzenie_kotwiczki ()

$t1 = 20000
;~ t1 - czas �adowania okna znaku wype�niacza (8000 dla max 10 zmiennych
$t2 = 3000
;~ t2 - czas ukazywania si� wszystkich w�a�ciwo�ci i w�a�ciwo�ci strony w oknie znaku wype�niacza po ich zafajkowaniu
$tc = 2000
;~ tc - copy time


;~ 1. Skopiowanie nazw pe�nych do excela
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

;~ 2. Usuni�cie nazw pe�nych w excelu
WinActivate("Usuwanie_nazw_pe�nych")
   Send("^{HOME}")
   Send("^{v}")
   Send("+{Enter}")
   Send("{RIGHT}")
   Send("^+{UP}")
   Send("+{DOWN}")
   Send("^{c}")
   Sleep($tc)

;~ 3. Umieszczenie nazw wy�wietlanych i kopiowanie w�a�ciwo�ci
WinActivate("W�a�ciwo�ci (symbol graficzny)")
   Send("{RIGHT}")
   Send("{down}")
   Send("^{v}")
   Send("^{TAB}")
   Sleep($tc)
   Send("^{a}")
   Sleep($tc)
   Send("^{c}")
   Sleep($tc)

;~ 4. Zamiana W�a�ciwo�ci na prawid�owe w excelu
WinActivate("Usuwanie_nazw_pe�nych")
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

;~ 5. Stworzenie nowego obiektu w�a�ciwo�ci i przypo��dkowanie prawid�owych zmiennych
WinActivate("W�a�ciwo�ci (symbol graficzny)")
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