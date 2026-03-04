' _lanzar_servidor.vbs — Lanza NEXUS sin ventana negra
Dim objShell, objFSO, carpeta, pythonCmd, configPath, f, linea
Set objShell = CreateObject("WScript.Shell")
Set objFSO   = CreateObject("Scripting.FileSystemObject")
carpeta    = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
configPath = carpeta & ".nexus_config"
pythonCmd  = "python"
If objFSO.FileExists(configPath) Then
    Set f = objFSO.OpenTextFile(configPath, 1)
    Do While Not f.AtEndOfStream
        linea = f.ReadLine()
        If Left(linea, 11) = "python_cmd=" Then
            pythonCmd = Mid(linea, 12)
        End If
    Loop
    f.Close
End If
objShell.CurrentDirectory = carpeta
objShell.Run Chr(34) & pythonCmd & Chr(34) & " nexus_tray.py", 0, False
Set objShell = Nothing
Set objFSO   = Nothing
