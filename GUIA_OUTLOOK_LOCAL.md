# Envio directo desde Outlook clasico - version basada en el flujo de borradores que funcionaba

Esta revision conserva la conexion COM, seleccion de cuenta, validacion y creacion del mensaje del script original que generaba borradores. La operacion `Save()` ahora es `Send()`.

1. Actualice GitHub con todos los archivos y carpetas de este ZIP, conservando las rutas.
2. Desde el gestor genere el paquete de constancias con QR.
3. Extraiga todo el paquete de constancias; abra Outlook clasico en Windows.
4. Ejecute `INICIAR_ENVIO.cmd` en la carpeta extraida.
5. Escriba una cuenta que aparezca en Outlook como cuenta remitente (si CENDOC solo es un buzon adicional, puede no aparecer).
6. Tras validar el total, escriba exactamente `ENVIAR` para emitir los mensajes.
7. Revise `REPORTE_ENVIO.csv`, Elementos enviados y Bandeja de salida. El reporte indica que Outlook acepto la orden, NO que haya llegado a destino.
8. No vuelva a ejecutar el lote antes de revisar que no haya mensajes entregados o pendientes: podria producir duplicados.

IMPORTANTE: pruebe con UN SOLO destinatario propio antes de usar el lote completo. Los envios reales requieren permisos de la cuenta y Outlook conectado. No incluye capacidades de envio desde buzon compartido sin los permisos de Exchange. Agregar QR despues de la firma puede invalidar la firma digital.
