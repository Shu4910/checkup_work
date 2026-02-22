Dim objShell
Set objShell = CreateObject("WScript.Shell")

Dim projectDir
projectDir = "C:\Users\sz83h\OneDrive\python_projects\checkup_work"

Dim pythonw
pythonw = "C:\Users\sz83h\AppData\Local\Python\bin\pythonw.exe"

Dim script
script = projectDir & "\src\activity_monitor.py"

objShell.CurrentDirectory = projectDir
objShell.Run """" & pythonw & """ """ & script & """", 0, False

Set objShell = Nothing
