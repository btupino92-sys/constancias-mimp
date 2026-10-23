# Módulo integral: Outlook web, nuevo y clásico

## Qué es posible desde GitHub Pages

- **Outlook web, nuevo Outlook y Outlook clásico con Exchange Online:** se usa la misma API Microsoft Graph, independientemente de qué cliente abra después los elementos enviados. Requiere registro SPA (URL de GitHub Pages como redirect URI), identificador de aplicación y directorio de Microsoft Entra, consentimiento delegado `Mail.Send` y `Mail.Send.Shared` cuando se use buzón compartido. También requiere permiso **Send As** o **Send on Behalf** de `cendoc@mimp.gob.pe`, gestionado por OGTI.
- **Outlook clásico local que no usa Exchange Online:** botón de descarga de paquete. Descomprimir ZIP, abrir Excel de escritorio e importar `outlook/PrepararBorradoresConstancias.bas`. Se crean borradores para revisión; no se puede abrir Outlook COM directamente desde una web hospedada.
- **Nuevo Outlook sin Graph/Exchange Online:** no tiene la automatización COM de Outlook clásico. Es necesario que OGTI habilite un envío mediante API/servicio autorizado; no se promete envío directo.

## Instrucciones

1. Completa los pasos 1 a 4 y comprueba asociaciones del PDF y el enlace.
2. Para Microsoft 365 selecciona el modo en paso 5, escribe los GUID proporcionados por OGTI y conecta tu sesión. Prepara la lista; verifica remitente, destinatarios y asunto; haz primero un envío a una dirección de prueba. Confirma el envío real de los seleccionados. Descarga el reporte.
3. Para Outlook clásico local selecciona modo local y usa «Descargar paquete para Outlook clásico». No necesitas hacer carga adicional: usa los mismos datos registrados.
4. No publiques Excel de participantes, PDF firmados, tokens, contraseñas o secretos en GitHub.

## Límites

- PDF adjunto de hasta 2,8 MiB en esta versión de envío Graph; para archivos mayores se requiere un flujo de subida por partes con Graph.
- Los navegadores pueden bloquear ventanas emergentes de autenticación; habilítalas solo para el dominio publicado.
- La aceptación por Graph no equivale a una entrega confirmada al destinatario.
- No se almacena historial de envío entre recargas; exporta el reporte antes de cerrar la pestaña.
- Insertar QR después de la firma puede invalidar la firma digital; el paso 4 conserva copia intacta.
- Para cargas masivas usar lotes limitados y seguir las políticas institucionales de mensajería y seguridad.
