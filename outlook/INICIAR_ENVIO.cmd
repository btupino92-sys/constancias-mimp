@echo off
setlocal
cd /d "%~dp0"
title Gestor MIMP - Envio de constancias
if not exist "ENVIO_OUTLOOK_LOCAL.csv" (
 echo ERROR: No se encontro ENVIO_OUTLOOK_LOCAL.csv en esta carpeta.
 echo Extraiga el ZIP completo y ejecute el CMD desde la carpeta extraida.
 pause
 exit /b 1
)
echo ATENCION: Este programa enviara mensajes reales al confirmar ENVIAR.
echo Revise Elementos enviados y Bandeja de salida ANTES de repetir.
powershell.exe -NoProfile -File "%~dp0EnviarConstanciasOutlookClasico.ps1"
echo.
echo Revise REPORTE_ENVIO.csv y DIAGNOSTICO_ENVIO_*.txt si se crearon.
echo La ventana permanecera abierta para consultar errores.
pause
