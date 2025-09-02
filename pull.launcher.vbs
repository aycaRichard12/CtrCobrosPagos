If WScript.Arguments.Count > 0 Then
    projectPath = WScript.Arguments(0)
    Set WshShell = CreateObject("WScript.Shell")
    WshShell.Run "cmd /c cd /d """ & projectPath & """ && git pull origin main", 0, True
Else
    MsgBox "Error: No se recibió la ruta del proyecto."
End If