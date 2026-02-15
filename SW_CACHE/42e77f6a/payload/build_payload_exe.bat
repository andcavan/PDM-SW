@echo off
setlocal
REM Build EXE payload (richiede pyinstaller)
REM Eseguire da questa cartella: WORKSPACES\<ws_id>\macros\payload

REM Usa python della venv del PDM se esiste
set PY="C:\PDM-SW\.venv\Scripts\python.exe"
if not exist %PY% set PY=python

%PY% -m pip show pyinstaller >nul 2>nul
if errorlevel 1 (
  echo PyInstaller non trovato. Installa con:
  echo   %PY% -m pip install pyinstaller
  pause
  exit /b 1
)

%PY% -m PyInstaller --noconfirm --onefile --windowed --name PDM_SW_PAYLOAD "PDM_SW_PAYLOAD.py"
if errorlevel 1 (
  echo Build fallita.
  pause
  exit /b 1
)

echo OK. EXE creato in dist\PDM_SW_PAYLOAD.exe
echo Copia PDM_SW_PAYLOAD.exe nella stessa cartella del payload (qui).
copy /Y "dist\PDM_SW_PAYLOAD.exe" ".\PDM_SW_PAYLOAD.exe" >nul
echo Fatto.
pause
