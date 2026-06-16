@echo off
cd /d C:\Robo_eGestor
echo [%date% %time%] Iniciando o Robo e-Gestor...
python automacao_sistema.py
if %errorlevel% neq 0 (
    echo [%date% %time%] ERRO: O robo terminou com falha. Verifique logs\robo_egestor.log
)
pause