@echo off
:: Executa o script de forma invisível passando o parametro "run"
if "%~1"=="run" goto :do_run

:: Cria script VBS temporário para chamar o próprio .bat ocultando a tela preta
set "vbsFile=%temp%\run_hidden_numerador.vbs"
echo Set objShell = CreateObject("WScript.Shell") > "%vbsFile%"
echo objShell.Run "cmd /c """"%~f0"" run""", 0, False >> "%vbsFile%"
cscript //nologo "%vbsFile%"
del "%vbsFile%"
exit /b

:do_run
:: Muda para o diretorio onde o arquivo .bat esta localizado
cd /d "%~dp0"

:: Executa o python silencioso (sem terminal) da pasta runtime e fecha
start "" "runtime\python\pythonw.exe" app\main.py
exit /b
