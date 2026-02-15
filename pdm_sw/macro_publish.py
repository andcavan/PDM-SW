from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Tuple
from datetime import datetime
from .workspace import WorkspaceManager

# Usiamo triple-apici singoli per includere liberamente virgolette VBA.

BOOTSTRAP_BAS_TEMPLATE = r'''Attribute VB_Name = "PDM_SW_BOOTSTRAP"
Option Explicit

' ===== PDM SolidWorks Macro Bootstrap =====
' Workspace bloccata (non selezionabile): WS_ID (cartella: WS_FOLDER)
' Questo bootstrap copia il payload in cache locale e lo lancia.
'
' Generato dal PDM.

Const PDM_ROOT As String = "{PDM_ROOT}"
Const WS_ID As String = "{WS_ID}"
Const WS_FOLDER As String = "{WS_FOLDER}"  ' cartella workspace (es. id_nome)

Const PY_EXE As String = "{PY_EXE}"  ' python.exe o pythonw.exe usato dal PDM

Const PAYLOAD_EXE As String = "PDM_SW_PAYLOAD.exe"
Const PAYLOAD_PY As String = "PDM_SW_PAYLOAD.py"

Const DEBUG_MODE As Boolean = True

Private Function JsonEscape(ByVal s As String) As String
    ' Escape minimo per JSON in riga comando: \ e "
    s = Replace(s, "", "")
    s = Replace(s, Chr(34), "" & Chr(34))
    s = Replace(s, vbCrLf, "\\n")
    s = Replace(s, vbLf, "\\n")
    JsonEscape = s
End Function

Private Sub EnsureDir(ByVal p As String)
    On Error Resume Next
    If Len(Dir(p, vbDirectory)) = 0 Then
        MkDir p
    End If
End Sub

Private Sub AppendLog(ByVal logPath As String, ByVal msg As String)
    On Error Resume Next
    Dim fso As Object, ts As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(logPath, 8, True) ' ForAppending = 8
    ts.WriteLine Now & " | " & msg
    ts.Close
End Sub

Private Sub CopyFolderOverwrite(ByVal src As String, ByVal dst As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    If Not fso.FolderExists(dst) Then
        fso.CreateFolder dst
    End If
    fso.CopyFile src & "\\*", dst & "", True
End Sub

Sub main()
    On Error GoTo EH

    Dim q As String
    q = Chr(34)

    ' Late binding: evita problemi di riferimenti VBA
    Dim swApp As Object
    Set swApp = Application.SldWorks

    Dim doc As Object
    Set doc = Nothing
    On Error Resume Next
    Set doc = swApp.IActiveDoc2
    If doc Is Nothing Then
        Set doc = swApp.ActiveDoc
    End If
    On Error GoTo EH

    Dim pathName As String
    pathName = ""
    If Not doc Is Nothing Then
        On Error Resume Next
        pathName = doc.GetPathName
        On Error GoTo EH
    End If

    ' Source payload (published by PDM)
    Dim srcPayload As String
    srcPayload = PDM_ROOT & "\\WORKSPACES\\" & WS_FOLDER & "\\macros\\payload"

    ' Cache payload
    Dim cacheBase As String
    cacheBase = PDM_ROOT & "\\SW_CACHE\\" & WS_ID & "\\payload"
    EnsureDir PDM_ROOT & "\\SW_CACHE"
    EnsureDir PDM_ROOT & "\\SW_CACHE\\" & WS_ID
    EnsureDir cacheBase

    Dim logPath As String
    logPath = cacheBase & "\\bootstrap.log"
    AppendLog logPath, "Bootstrap start. ActivePath=" & pathName

    ' Copy payload (best-effort, overwrite)
    CopyFolderOverwrite srcPayload, cacheBase
    AppendLog logPath, "Payload copied from " & srcPayload & " to " & cacheBase

' Build sw-context JSON file (evita problemi di quote in riga comando)
Dim ctxPath As String
ctxPath = cacheBase & "\sw_context.json"

Dim fso As Object, ts As Object
Set fso = CreateObject("Scripting.FileSystemObject")
Set ts = fso.OpenTextFile(ctxPath, 2, True) ' ForWriting = 2, create=True
Dim swpid As Long
swpid = 0
On Error Resume Next
swpid = swApp.GetProcessID
On Error GoTo 0
ts.Write "{{" & q & "active_doc_path" & q & ":" & q & JsonEscape(pathName) & q & "," & q & "sw_pid" & q & ":" & CStr(swpid) & "}}"
ts.Close
AppendLog logPath, "SW context file: " & ctxPath

Dim exePath As String
    exePath = cacheBase & "\\" & PAYLOAD_EXE

    Dim payloadLog As String
    payloadLog = cacheBase & "\payload.log"
    Dim batPath As String
    batPath = cacheBase & "\run_payload.bat"

    Dim runner As String
    runner = ""
    If Len(Dir(exePath)) > 0 Then
        runner = q & exePath & q & " --pdm-root " & q & PDM_ROOT & q & " --workspace " & WS_ID & " --sw-context-file " & q & ctxPath & q & " --log-file " & q & payloadLog & q
        AppendLog logPath, "Use EXE: " & exePath
    Else
        Dim pyw As String
        pyw = PY_EXE
        If Len(Dir(pyw)) = 0 Then
            pyw = "pythonw"
        End If
        runner = q & pyw & q & " " & q & cacheBase & "\\" & PAYLOAD_PY & q & " --pdm-root " & q & PDM_ROOT & q & " --workspace " & WS_ID & " --sw-context-file " & q & ctxPath & q & " --log-file " & q & payloadLog & q
        AppendLog logPath, "Use PY: " & pyw
    End If

    Dim fsoR As Object, tsR As Object
    Set fsoR = CreateObject("Scripting.FileSystemObject")
    Set tsR = fsoR.CreateTextFile(batPath, True, False)
    tsR.WriteLine "@echo off"
    tsR.WriteLine "setlocal"
    tsR.WriteLine "cd /d " & q & cacheBase & q
    tsR.WriteLine "echo ==== START %DATE% %TIME% ====>>" & q & payloadLog & q
    tsR.WriteLine runner & " >> " & q & payloadLog & q & " 2>&1"
    tsR.WriteLine "echo ==== EXIT %ERRORLEVEL% %DATE% %TIME% ====>>" & q & payloadLog & q
    tsR.WriteLine "endlocal"
    tsR.Close

    Dim cmd As String
    cmd = "cmd.exe /c " & q & q & batPath & q & q
    AppendLog logPath, "Run BAT: " & cmd
    Dim pid As Double
    pid = Shell(cmd, vbHide)
    AppendLog logPath, "Shell PID: " & CStr(pid)

    Exit Sub

EH:
    On Error Resume Next
    Dim cacheBaseEH As String, logPathEH As String
    cacheBaseEH = PDM_ROOT & "\\SW_CACHE\\" & WS_ID & "\\payload"
    logPathEH = cacheBaseEH & "\\bootstrap.log"
    AppendLog logPathEH, "ERROR " & Err.Number & " - " & Err.Description
    MsgBox "Errore macro PDM bootstrap: " & Err.Description & vbCrLf & "Vedi log: " & logPathEH, vbExclamation, "PDM"
End Sub
'''

PAYLOAD_PY_TEMPLATE = r'''# -*- coding: utf-8 -*-
# PDM SolidWorks Payload (launcher)
# Generato dal PDM - workspace: {WS_ID}
#
# Questo file lancia la UI minimale per Codifica/Workflow da SolidWorks.

import sys
from pathlib import Path

def _extract_pdm_root(argv):
    try:
        i = argv.index("--pdm-root")
        return argv[i + 1]
    except Exception:
        return ""

root = _extract_pdm_root(sys.argv)
if root:
    pdm_root = Path(root)
    if str(pdm_root) not in sys.path:
        sys.path.insert(0, str(pdm_root))

from pdm_sw.macro_runtime import main  # noqa

if __name__ == "__main__":
    main()
'''

BUILD_BAT_TEMPLATE = r'''@echo off
setlocal
REM Build EXE payload (richiede pyinstaller)
REM Eseguire da questa cartella: WORKSPACES\<ws_id>\macros\payload

REM Usa python della venv del PDM se esiste
set PY="{PDM_ROOT}\.venv\Scripts\python.exe"
if not exist %PY% set PY=python

%PY% -m pip show pyinstaller >nul 2>nul
if errorlevel 1 (
  echo PyInstaller non trovato. Installa con:
  echo   %PY% -m pip install pyinstaller
  pause
  exit /b 1
)

%PY% -m PyInstaller --noconfirm --onefile --windowed --name PDM_SW_PAYLOAD "{PAYLOAD_PY}"
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
'''

INSTALL_TXT_TEMPLATE = r'''INSTALLAZIONE MACRO SOLIDWORKS (WORKSPACE {WS_ID})
====================================================

Obiettivo:
- Lanciare una finestra "PDM (Macro SolidWorks)" direttamente da SolidWorks,
  con WORKSPACE BLOCCATA (non selezionabile).
- Permette: Codifica (crea codice + SaveAs in WIP) e Workflow (WIP/REL/IN_REV/OBS).

Percorsi:
- Root PDM: {PDM_ROOT}
- Bootstrap sorgente: {BOOTSTRAP_SRC}
- Payload: {PAYLOAD_DIR}
- Cache: {PDM_ROOT}\SW_CACHE\{WS_ID}\payload

PASSI (una tantum):
1) In SolidWorks: Strumenti > Macro > Nuova...
2) Salva la macro come:
   {PDM_ROOT}\SW_MACROS\PDM_SW_BOOTSTRAP_{WS_ID}.swp
3) SolidWorks aprirà l'editor VBA. Importa il modulo:
   File > Import File...  e seleziona:
   {BOOTSTRAP_SRC}
4) Salva e chiudi l'editor VBA.
5) (Consigliato) Crea un pulsante:
   Strumenti > Personalizza... > Comandi > Macro
   trascina "Esegui Macro" su toolbar e seleziona il file .swp creato.

BUILD EXE (consigliato):
- Vai in:
  {PAYLOAD_DIR}
- Esegui:
  build_payload_exe.bat
- Otterrai:
  PDM_SW_PAYLOAD.exe
Il bootstrap userà l'EXE se presente, altrimenti userà pythonw.
'''

def _now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def publish_macro(pdm_root: Path, ws_id: str) -> Tuple[Path, Path]:
    """Pubblica bootstrap + payload in WORKSPACES/<ws_id>/macros e crea una copia in SW_MACROS.
    Ritorna (bootstrap_bas_path, payload_dir).
    """
    # Interpreter da usare per il payload quando viene lanciato dalla macro.
    # Usiamo lo stesso Python che sta eseguendo il PDM (così customtkinter & dipendenze ci sono).
    py_exe = sys.executable
    try:
        # preferisci pythonw.exe se esiste accanto a python.exe (evita console popup)
        p = Path(py_exe)
        if p.name.lower() == 'python.exe':
            cand = p.with_name('pythonw.exe')
            if cand.exists():
                py_exe = str(cand)
    except Exception:
        pass
    pdm_root = Path(pdm_root)
    # Risolvi cartella reale della workspace da WORKSPACES/workspaces.json (id -> path).
    # Le workspace sono cartelle "id_nome" (es. 8564ba90_ClienteX).
    try:
        wm = WorkspaceManager(pdm_root / "WORKSPACES")
        ws_dir = wm.workspace_dir(ws_id)
    except Exception:
        ws_dir = (pdm_root / "WORKSPACES" / ws_id)
    ws_folder = ws_dir.name
    macros_dir = ws_dir / "macros"
    payload_dir = macros_dir / "payload"
    bootstrap_dir = macros_dir / "bootstrap"

    payload_dir.mkdir(parents=True, exist_ok=True)
    bootstrap_dir.mkdir(parents=True, exist_ok=True)

    payload_json = {
        "schema": 1,
        "ws_id": ws_id,
        "name": "PDM SolidWorks Macro Payload",
        "version": "0.1.0",
        "built_at": _now_iso(),
        "entry_exe": "PDM_SW_PAYLOAD.exe",
        "entry_py": "PDM_SW_PAYLOAD.py",
        "notes": "Payload UI minimale (Codifica/Workflow) lanciato da macro SolidWorks."
    }
    (payload_dir / "payload.json").write_text(json.dumps(payload_json, indent=2, ensure_ascii=False), encoding="utf-8")

    (payload_dir / "PDM_SW_PAYLOAD.py").write_text(
        PAYLOAD_PY_TEMPLATE.format(WS_ID=ws_id),
        encoding="utf-8"
    )

    (payload_dir / "build_payload_exe.bat").write_text(
        BUILD_BAT_TEMPLATE.format(PDM_ROOT=str(pdm_root), PAYLOAD_PY="PDM_SW_PAYLOAD.py"),
        encoding="utf-8"
    )

    bas_path = bootstrap_dir / f"PDM_SW_BOOTSTRAP_{ws_id}.bas"
    bas_path.write_text(
        BOOTSTRAP_BAS_TEMPLATE.format(PDM_ROOT=str(pdm_root), WS_ID=ws_id, WS_FOLDER=ws_folder, PY_EXE=py_exe),
        encoding="utf-8"
    )

    sw_macros = pdm_root / "SW_MACROS"
    sw_macros.mkdir(parents=True, exist_ok=True)
    bas_copy = sw_macros / f"PDM_SW_BOOTSTRAP_{ws_id}.bas"
    bas_copy.write_text(bas_path.read_text(encoding="utf-8"), encoding="utf-8")

    install_txt = sw_macros / f"INSTALL_MACRO_{ws_id}.txt"
    install_txt.write_text(
        INSTALL_TXT_TEMPLATE.format(
            WS_ID=ws_id,
            PDM_ROOT=str(pdm_root),
            BOOTSTRAP_SRC=str(bas_copy),
            PAYLOAD_DIR=str(payload_dir)
        ),
        encoding="utf-8"
    )

    return bas_copy, payload_dir
