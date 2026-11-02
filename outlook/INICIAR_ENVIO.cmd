@echo off
cd /d "%~dp0"
echo Gestor MIMP - envio directo de constancias por Outlook clasico
echo Se solicitara confirmacion ENVIAR antes de enviar.
powershell.exe -NoProfile -File "%~dp0EnviarConstanciasOutlookClasico.ps1"
echo.
echo Fin del proceso. Revise REPORTE_ENVIO.csv en esta carpeta.
pause
