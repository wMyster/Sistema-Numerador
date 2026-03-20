@echo off
:: ============================================================
:: DEBUG.BAT - Executa o sistema COM CONSOLE VISIVEL para testes
:: Mostra todos os logs, erros e prints diretamente na tela.
:: Use este arquivo para depurar problemas no sistema.
:: ============================================================
cd /d "%~dp0"

echo ============================================================
echo   SISTEMA NUMERADOR - MODO DEPURACAO (DEBUG)
echo   Todos os logs serao exibidos nesta janela.
echo   Feche esta janela para encerrar o sistema.
echo ============================================================
echo.

:: Usa python.exe (com console) ao inves de pythonw.exe (sem console)
if exist "runtime\python\python.exe" (
    "runtime\python\python.exe" app\main.py --debug
) else (
    python app\main.py --debug
)

echo.
echo ============================================================
echo   Sistema encerrado. Pressione qualquer tecla para fechar.
echo ============================================================
pause
