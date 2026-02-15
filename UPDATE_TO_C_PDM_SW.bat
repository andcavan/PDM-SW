@echo off
set TARGET=C:\PDM-SW
echo.
echo Aggiornamento PDM-SW -> %TARGET%
echo (NON tocca la cartella WORKSPACES del target, se non presente nello zip)
echo.
REM Copia tutti i file del pacchetto nel target.
xcopy /E /I /Y /Q "%~dp0*" "%TARGET%\"
echo.
echo FATTO. Ora avvia: python %TARGET%\app.py
pause
