# Paso 5 - envío directo desde Outlook clásico de Windows

1. Completa participantes, correos, enlaces y PDF firmados; verifica las asociaciones.
2. Selecciona **Outlook clásico Windows — envío directo sin Excel** y descarga el paquete.
3. Extrae TODO el ZIP. Confirma que estén `ENVIO_OUTLOOK_LOCAL.csv`, `DIPLOMAS_FINALES_QR`, `INICIAR_ENVIO.cmd` y `EnviarConstanciasOutlookClasico.ps1`.
4. Abre Outlook clásico en Windows y verifica la conexión. Ejecuta `INICIAR_ENVIO.cmd`.
5. Introduce una dirección de cuenta que figure como cuenta remitente en Outlook. El sistema valida correos y PDF.
6. Escribe exactamente **ENVIAR** para confirmar el envío real. La ejecución solicita el envío de cada mensaje con su PDF adjunto, sin guardar borradores.
7. Revisa el archivo `REPORTE_ENVIO_*.csv`, **Elementos enviados** y **Bandeja de salida**. La orden aceptada por Outlook no garantiza la entrega final.

**IMPORTANTE:** ejecutar el mismo paquete otra vez puede duplicar los correos; comprueba el reporte antes de repetir. No subas información personal o PDFs generados a GitHub. El método exige Outlook clásico de Windows, no funciona con Outlook nuevo o web y depende de permisos institucionales. No eludas políticas de OGTI. Agregar QR a un PDF firmado puede invalidar su firma digital.
