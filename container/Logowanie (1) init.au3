Login_user()
$Program_name = 'Login_test'

Func Login_user()

	Global $login_me[2], $login_tomek[2], $login_pawel, $login_lukasz, $login_Ola
	Local $Txt_login, $login_wrong, $login_restart, $password_me, $password_glibal
	$login_me[0] = "no"
	$login_me[1] = "danditkaczuk"
	$login_tomek[0] = "123"
	$login_tomek[1] = "glinoconto"
	$login_pawel = "glikozowpa"
	$login_lukasz = "gliwoliclu"
	$login_Ola = "ola"
	$Txt_login = "gli pc_user_login"
	$Txt_password = 'puser_assword'
	$password_me = ' '
	$password_glibal = 'walmed'
	$login = InputBox($Program_name, "Prosze, wpisz swój 10 literowy login", $Txt_login)
	$password = InputBox("Tworzenie kotwiczki z Darkiem :D", "Prosze, wpisz swóje haslo", $Txt_password)
	If $login = $login_me[0] Or $login = $login_me[1] Then
		$login = "glitkaczda"
	ElseIf $login = $login_tomek[0] Or $login = $login_tomek[1] Then
		$login = "glinoconto"
		MsgBox(0, $Program_name, "Czesc Tomasz ;D ", 2)
	ElseIf $login = $login_pawel Then
		$login = "glikozowpa"
		MsgBox(0, $Program_name, "Czesc Pewel !", 2)
	ElseIf $login = $login_lukasz Then
		$login = "gliwoliclu"
		MsgBox(0, $Program_name, "Czesc Lukasz !", 2)
	ElseIf $login = $login_Ola Then
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
