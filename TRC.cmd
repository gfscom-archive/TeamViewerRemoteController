::	TeamViewer Remote Controller
::	-----------------------------------------------------------------------------------------------------------
::	Autore:				GSolone
::	Versione:			0.4
::	Utilizzo:			.\TRC.cmd NOMEMACCHINA
::						ATTENZIONE: va utilizzato come amministratore locale o 
::						di dominio della macchina alla quale si punta.
::	Info:				https://gioxx.org/tag/teamviewer
::	Ultima modifica:	07-02-2018
::	Fonti utilizzate:	https://gist.github.com/mlocati/fdabcaeb8071d5c75a2d51712db24011#file-win10colors-cmd
::						http://www.robvanderwoude.com/parameters.php
::						https://stackoverflow.com/questions/4419868/what-is-the-current-directory-in-a-batch-file
::	------------------------------------------------------------------------------------------------------------
::	Modifiche:
::		0.4- evito la ripartenza del ciclo StartEngine al termine di una modifica via prompt / TRC.
::		0.3- ho isolato lo script e tutti i VBS, per poterlo "trasportare facilmente altrove". Ho integrato la nuova funzione di modifica Permanent Password da remoto. Ho tolto anche i riferimenti assoluti al programma.
::		0.2- modificate spaziature in output VBS, modificato colore dell'output.

@echo off
cls

if not exist %~dp0bin goto BinMancante

IF "%1" == "" GOTO DatiMancanti
IF "%1" NEQ "" GOTO StartEngine

goto END

:StartEngine
echo.
echo   *******************************************************
echo     TeamViewer Remote Controller
echo   *******************************************************
echo.
echo    [93mMacchina da controllare: %1 [0m
echo.
echo    (0)  Esci dal programma
echo    (1)  REMOTE KILL / START
echo    (2)  REMOTE PASSWORD
echo    (3)  REMOTE HOST
echo    (4)  REMOTE HOST + PASSWORD
echo.
SET /P SCELTA="Seleziona opzione (es. 1): "
REM echo DEBUG %errorlevel%

if errorlevel 1 set "SCELTA=" & verify>nul & goto StartEngine
IF /i %SCELTA% EQU 0 goto END
IF /i %SCELTA% EQU 1 goto RemoteKill
IF /i %SCELTA% EQU 2 goto RemotePassword
IF /i %SCELTA% EQU 3 goto RemoteHost
IF /i %SCELTA% EQU 4 goto HostPassword

echo;
goto END

:RemoteKill
echo [92m
cscript /nologo %~dp0bin\TeamViewerRemoteKill.vbs %1
echo [0m
goto END

:RemotePassword
echo [92m
cscript /nologo %~dp0bin\TeamViewerRemotePasswd.vbs %1
echo [0m
goto END

:RemoteHost
echo [92m
cscript /nologo %~dp0bin\TeamViewerRemoteHost.vbs %1
echo [0m
goto END

:HostPassword
echo [92m
cscript /nologo %~dp0bin\TeamViewerRemotePasswd.vbs %1
cscript /nologo %~dp0bin\TeamViewerRemoteHost.vbs %1
echo [0m
goto END

::	--  Messaggi di programma  ---------------------------------------------------------------------------------

:DatiMancanti
echo.
echo [93mATTENZIONE[0m
echo [93mMacchina da controllare non specificata.[0m
echo [93mRilancia lo script con i giusti parametri.[0m
echo Esempi:
echo  - %~dpnx0 127.0.0.1
echo  - %~dpnx0 SERVER01.localdomain.local
echo.
goto END

:BinMancante
echo.
echo [93mATTENZIONE[0m
echo [93mSembra che la cartella bin non sia presente, check file:[0m
if exist %~dp0bin\TeamViewerRemoteKill.vbs (echo  - %~dp0bin\TeamViewerRemoteKill.vbs trovato!) else (echo  - %~dp0bin\TeamViewerRemoteKill.vbs NON trovato!)
if exist %~dp0bin\TeamViewerRemotePasswd.vbs (echo  - %~dp0bin\TeamViewerRemotePasswd.vbs trovato!) else (echo  - %~dp0bin\TeamViewerRemotePasswd.vbs NON trovato!)
if exist %~dp0bin\TeamViewerRemoteHost.vbs (echo  - %~dp0bin\TeamViewerRemoteHost.vbs trovato!) else (echo  - %~dp0bin\TeamViewerRemoteHost.vbs NON trovato!)
echo.
echo Scarica i file necessari e integrali all'interno della cartella del programma (%~dp0bin\)
echo.

:END