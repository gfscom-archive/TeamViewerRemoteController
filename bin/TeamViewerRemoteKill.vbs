' TeamViewer Remote Kill for TeamViewer 10+
' GSolone 2015 v 0.2
' Ultima modifica 10012018
'
' Funzione di ricerca processi: http://www.freevbcode.com/ShowCode.asp?ID=4888
'
' LO SCRIPT VA ESEGUITO COME AMMINISTRATORE MACCHINA O DI DOMINIO!

Set WSHShell = WScript.CreateObject("WScript.Shell")
Set WNetwork = Wscript.CreateObject("Wscript.Network")
On Error Resume Next

Function IsProcessRunning( strServer, strProcess )
    Dim Process, strObject
    IsProcessRunning = False
    strObject   = "winmgmts://" & strServer
    For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
	If UCase( Process.name ) = UCase( strProcess ) Then
            IsProcessRunning = True
            Exit Function
        End If
    Next
End Function

if WScript.Arguments.Count = 0 then
    'Richiesta IP o hostname macchina da modificare
	strComputer = InputBox("TeamViewer Remote Kill" & vbCR & "Lo script permette di terminare da remoto TeamViewer, sia per x86 che x64, rileva automaticamente l'installazione." & vbCR & vbCR & "VA ESEGUITO COME ADMIN LOCALI O DI DOMINIO!" & vbCR & vbCR & "Hostname macchina di destinazione (vuoto o clic su Annulla per uscire dallo script)" & vbCR, "TV Remote Kill", "COMPUTERNAME")
else
	'Se l'indirizzo IP / Hostname mi è stato passato da riga di comando, posso procedere direttamente
	strComputer = Wscript.Arguments(0)
end if

if LEN(trim(strComputer)) = 0 Then
	Messaggio = msgbox ("Non hai indicato l'hostname di destinazione, termino lo script adesso.", vbCritical, "")
	wscript.quit
else
	
	'Controllo stato del TeamViewer
	strProcess = "TeamViewer_Service.exe"
	if( IsProcessRunning( strComputer, strProcess ) = True ) Then
		
		'Richiedo possibilità di terminare
		WScript.Echo strProcess & " avviato (" & strComputer & ")"
		strQuestion = "Devo terminare TeamViewer sulla macchina remota?"
		answer= msgbox (strQuestion, vbYesNo)
		if answer= vbYes then 
			'Termino TV
			WshShell.Run "sc \\" & strComputer & " stop Teamviewer"
			WshShell.Run "taskkill /s " & strComputer & " /IM TeamViewer_Service.exe /F"
			WshShell.Run "taskkill /s " & strComputer & " /IM TeamViewer.exe /F"
		end if
		
	else
	
		'TV non avviato
		WScript.Echo strProcess & " NON avviato (" & strComputer & ")"
		strQuestion = "Devo avviare TeamViewer sulla macchina remota?"
		answer= msgbox (strQuestion, vbYesNo)
		if answer= vbYes then 
			'Avvio TV
			WshShell.Run "sc \\" & strComputer & " start Teamviewer"
		end if
	
	end If

end if