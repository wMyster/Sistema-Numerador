Set objShell = CreateObject("WScript.Shell")
' Executando com o parametro 0, a janela do prompt nao sera exibida
objShell.Run "cmd /c """ & CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\run.bat""", 0
Set objShell = Nothing
