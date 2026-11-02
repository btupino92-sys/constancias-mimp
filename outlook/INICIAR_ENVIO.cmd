@echo off
cd /d "%~dp0"
echo Gestor de Constancias MIMP - ENVIO DIRECTO Outlook clasico
echo Abra Outlook clasico antes de continuar.
powershell.exe -NoProfile -File "%~dp0EnviarConstanciasOutlookClasico.ps1"
echo.
pause
