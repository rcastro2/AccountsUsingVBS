'Use cscript to be able to echo to the console
WScript.Echo "Script Started"

Set shell = WScript.CreateObject ("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")
Set file = fso.OpenTextFile("computerAccounts.csv")

station = InputBox("Enter Computer Station")

Do While file.AtEndOfStream <> True
    arr = Split(file.ReadLine, ",")
    If arr(0) = station Then
    	username = "S" & arr(0) & "_PD" & arr(1)
        password = arr(2)
    	shell.run "cmd.exe /C net user " & username & " " & password & " /add"
        'shell.run "cmd.exe /C net user " & username & " /delete"
    End If

Loop

WScript.Echo "Script Ended"