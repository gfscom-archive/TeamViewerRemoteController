' TeamViewer Unattended Switch (TVHost) for TeamViewer 10+
' GSolone 2015 v 0.3
' Ultima modifica 10012018
'
' LO SCRIPT VA ESEGUITO COME AMMINISTRATORE MACCHINA O DI DOMINIO!

Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set WNetwork = Wscript.CreateObject("Wscript.Network")
On Error Resume Next

if WScript.Arguments.Count = 0 then
    'Richiesta IP o hostname macchina da modificare
	strComputer = InputBox("TeamViewer Host Switch" & vbCR & "Lo script permette di abilitare o disabilitare l'accesso host al TeamViewer, sia per x86 che x64, rileva automaticamente l'installazione." & vbCR & vbCR & "VA ESEGUITO COME ADMIN LOCALI O DI DOMINIO!" & vbCR & vbCR & "Hostname macchina di destinazione (vuoto o clic su Annulla per uscire dallo script)" & vbCR, "TV Remote Host Switch", "COMPUTERNAME")
else
	'Se l'indirizzo IP / Hostname mi è stato passato da riga di comando, posso procedere direttamente
	strComputer = Wscript.Arguments(0)
end if

if LEN(trim(strComputer)) = 0 Then
	Messaggio = msgbox ("Non hai indicato l'hostname di destinazione, termino lo script adesso.", vbCritical, "")
	wscript.quit
else
	'Termino TV, modifico il valore di registro per generare o disabilitare la password locale, avvio TV
	WshShell.Run "taskkill /s " & strComputer & " /IM TeamViewer_Service.exe /F"
	WshShell.Run "taskkill /s " & strComputer & " /IM TeamViewer.exe /F"
	WshShell.Run "sc \\" & strComputer & " stop Teamviewer"
	
	Set StdOut = WScript.StdOut
	Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")
	
	'Rilevo sistema a 64 bit
	strKeyPath64 = "SOFTWARE\Wow6432Node\TeamViewer\AccessControl"
	strValueName64 = "AC_Server_AccessControlType"
	oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath64,strValueName64,dwValue64
	
	If IsEmpty(dwValue64) or IsNull(dwValue64) Then
		'Il sistema non è a 64 bit
		
		'DEBUG - Togliere commento se necessario
		'Wscript.Echo "Sembra che non si tratti di un sistema a 64 bit"
		strKeyPath32 = "SOFTWARE\TeamViewer\AccessControl"
		strValueName32 = "AC_Server_AccessControlType"
		oReg.GetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath32,strValueName32,dwValue32
		'DEBUG - Togliere commento se necessario
		'Wscript.Echo "Debug lettura (32 bit): " & dwValue32
		
		if dwValue32 = 3 Then
			dwValue32 = 0
			oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath32,strValueName32,dwValue32
			'DEBUG - Togliere commento se necessario
			'Wscript.Echo "Debug scrittura (32 bit): " & dwValue32
			Wscript.Echo "TeamViewer permette l'accesso senza conferma su " & strComputer &" (OS 32 bit)"
		else
			dwValue32 = 3
			oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath32,strValueName32,dwValue32
			'DEBUG - Togliere commento se necessario
			'Wscript.Echo "Debug scrittura (32 bit): " & dwValue32
			Wscript.Echo "TeamViewer NON permette l'accesso senza conferma su " & strComputer &" (OS 32 bit)"
		end if
	else
		'Il sistema è a 64 bit
		
		'DEBUG - Togliere commento se necessario
		'Wscript.Echo "Debug lettura (64 bit): " & dwValue64
		if dwValue64 = 3 Then
			dwValue64 = 0
			oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath64,strValueName64,dwValue64
			'DEBUG - Togliere commento se necessario
			'Wscript.Echo "Debug scrittura (64 bit): " & dwValue64
			Wscript.Echo "TeamViewer permette l'accesso senza conferma su " & strComputer &" (OS 64 bit)"
		else
			dwValue64 = 3
			oReg.SetDWORDValue HKEY_LOCAL_MACHINE,strKeyPath64,strValueName64,dwValue64
			'DEBUG - Togliere commento se necessario
			'Wscript.Echo "Debug scrittura (64 bit): " & dwValue64
			Wscript.Echo "TeamViewer NON permette l'accesso senza conferma su " & strComputer &" (OS 64 bit)"
		end if
	end if
	
	WshShell.Run "sc \\" & strComputer & " start Teamviewer"
end if