@echo off
:: Muda para o diretorio onde o arquivo .bat esta localizado
cd /d "%~dp0"

:: Executa o python mudo e seguro para nao assustar o Antivirus corporativo
start "" "runtime\python\pythonw.exe" app\main.py
exit
