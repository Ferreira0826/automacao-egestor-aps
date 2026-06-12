@echo off
cd /d C:\Robo_SISAB
echo [%date% %time%] Iniciando o Robo SISAB...
python automacao_sistema.py
if %errorlevel% neq 0 (
    echo [%date% %time%] ERRO: O robo terminou com falha. Verifique logs\robo_sisab.log
)
pause