' _instalar.vbs
' Ejecuta la instalación automática sin mostrar ventanas negras
Dim objShell
Set objShell = CreateObject("WScript.Shell")
Dim carpeta
carpeta = Left(WScript.ScriptFullName, InStrRev(WScript.ScriptFullName, "\"))
objShell.CurrentDirectory = carpeta
' 1 = ventana normal (para que el usuario vea el progreso de instalación)
objShell.Run "python nexus_setup.py", 1, True
Set objShell = Nothing
