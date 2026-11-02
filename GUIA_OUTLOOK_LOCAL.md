# Envío directo de constancias MIMP — versión de diagnóstico

1. Reemplace los archivos de GitHub por los del ZIP actualizado, conservando carpetas y nombres.
2. Abra el gestor actualizado y genere un paquete NUEVO para un solo destinatario de prueba.
3. Extraiga el ZIP del paquete completamente y abra Outlook clásico en Windows.
4. Ejecute `INICIAR_ENVIO.cmd` y escriba una cuenta que esté registrada como remitente en Outlook.
5. Si el CSV es antiguo y no contiene la actividad, el script solicitará el tipo y tema; escríbalos tal como aparecen en la constancia.
6. Revise el total y escriba **ENVIAR** para proceder. Solo entonces se llama a Outlook para enviar.
7. Consulte `REPORTE_ENVIO.csv` y `DIAGNOSTICO_ENVIO_*.txt`; compruebe **Elementos enviados** y **Bandeja de salida**. `ACEPTADO_POR_OUTLOOK` no prueba la recepción final.

**Evite reenviar el mismo lote sin comprobar su estado:** se pueden producir duplicados. Para remitentes que son buzones compartidos, se requieren permisos y que Outlook los reconozca adecuadamente. Este código no puede garantizar el envío sin una prueba real en Outlook clásico.
