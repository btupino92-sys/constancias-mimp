@echo off
cd /d "%~dp0"
echo Gestor de Constancias MIMP - Outlook clasico
echo Abrir Outlook clasico antes de continuar.
powershell.exe -NoProfile -File "%~dp0PrepararBorradoresOutlookClasico.ps1"
echo.
pause
