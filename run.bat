@echo off
:: Muda para o diretorio onde o arquivo .bat esta localizado
cd /d "%~dp0"

:: Executa o python silencioso (sem terminal) apartir do executavel w (pythonw.exe) da pasta runtime e fecha o prompt
start "" "runtime\python\pythonw.exe" app\main.py

exit
