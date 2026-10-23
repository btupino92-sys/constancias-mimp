# Paso 5 sin Microsoft 365 — Outlook clásico de Windows

1. En el HTML, complete participantes (con correos), enlaces y PDF firmados. Revise las asociaciones en paso 4.
2. En paso 5 seleccione «Outlook clásico Windows — borradores sin Excel» y pulse «Descargar paquete para Outlook clásico».
3. Extraiga TODO el ZIP en una carpeta local (no ejecute nada dentro del ZIP). Verifique `DIPLOMAS_FINALES_QR`, `ENVIO_OUTLOOK_LOCAL.csv`, `INICIAR_BORRADORES.cmd`, `PrepararBorradoresOutlookClasico.ps1`.
4. Abra **Outlook clásico** en Windows. En la carpeta extraída haga doble clic en `INICIAR_BORRADORES.cmd` (puede requerir autorización por las políticas de Windows/OGTI). Indique la cuenta real configurada en Outlook y escriba `CREAR` para confirmar borradores.
5. Revise los mensajes en Borradores antes de enviar. El auxiliar **no envía automáticamente**. Se produce `REPORTE_BORRADORES.csv`.

**Limitaciones:** el auxiliar local usa COM de Outlook clásico y **no funciona** con Nuevo Outlook ni Outlook web. Si Windows restringe scripts, consulte con OGTI; no desactive ni eluda las políticas institucionales. Para enviar desde esas otras versiones sin Microsoft 365/Graph se requiere un servicio autorizado por OGTI. Nunca suba PDFs, listas, paquetes de salida ni datos de participantes a GitHub. Para buzones compartidos como cendoc@mimp.gob.pe debe estar visible como cuenta enviable y contar con autorización. El código QR agregado tras firmar puede invalidar la firma digital.
