Attribute VB_Name = "PDM_SW_BOOTSTRAP"
Option Explicit

' ===== PDM SolidWorks Macro Bootstrap =====
' Workspace bloccata (non selezionabile): WS_ID (cartella: WS_FOLDER)
' Questo bootstrap copia il payload in cache locale e lo lancia.
'
' Generato dal PDM.

Const PDM_ROOT As String = "c:\PDM-SW"
Const WS_ID As String = "9a5aa5c7"
Const WS_FOLDER As String = "9a5aa5c7_PROG3DAC"  ' cartella workspace (es. id_nome)

Const PY_EXE As String = "C:\Users\prog3\AppData\Local\Programs\Python\Python311\pythonw.exe"  ' python.exe o pythonw.exe usato dal PDM

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
ts.Write "{" & q & "active_doc_path" & q & ":" & q & JsonEscape(pathName) & q & "," & q & "sw_pid" & q & ":" & CStr(swpid) & "}"
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
