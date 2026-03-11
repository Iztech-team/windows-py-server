' Baraka Printer Server - Silent Background Launcher
' Starts the server with NO visible window. Output goes to server.log.
Set objShell = CreateObject("WScript.Shell")
objShell.CurrentDirectory = "c:\baraka-printer-server-windows"
Set objFSO = CreateObject("Scripting.FileSystemObject")

pythonwPath = "C:\Python312\pythonw.exe"
pythonPath  = "C:\Python312\python.exe"
logFile     = "c:\baraka-printer-server-windows\server.log"

If objFSO.FileExists(pythonwPath) Then
    objShell.Run Chr(34) & pythonwPath & Chr(34) & " printer_server.py", 0, False
Else
    objShell.Run "cmd /c " & Chr(34) & pythonPath & Chr(34) & " printer_server.py > " & Chr(34) & logFile & Chr(34) & " 2>&1", 0, False
End If
