# Gestor de Constancias MIMP — Pasos 1 a 5

## Publicar en GitHub Pages
Suba **el contenido de esta carpeta** a la raíz del repositorio: `index.html`, `app.js`, `extension.js`, `styles.css`, `assets/`, `outlook/` y `README.md`.

En Settings > Pages, elija Deploy from branch, `main`, `/(root)`.

## Flujo
1. Configurar datos del evento.
2. Importar participantes desde Excel.
3. Generar PDF individuales, enviarlos a firma digital mediante el procedimiento autorizado.
4. Cargar todos los PDF firmados desde la computadora, importar Excel con columnas `Nombre completo` y `Enlace` (Drive). Emparejamiento por nombre normalizado. El sistema muestra los faltantes.
5. Generar paquete con originales firmados + hojas complementarias QR + control Excel + macro VBA para crear **borradores** en Outlook clásico de Windows.

## Importante sobre firma digital
Agregar contenido a un PDF **después de la firma** puede invalidar la firma. Por eso la aplicación **no edita** los PDFs firmados. Emite un segundo PDF individual que contiene el QR y el enlace de Drive. Si necesita el QR dentro del mismo diploma, debe incorporarlo **antes de que el director firme**.

## Outlook
El archivo ZIP generado en el paso 5 contiene `MACRO_OUTLOOK/PrepararBorradoresConstancias.bas` y las instrucciones `LEEME_ENVIO.txt`. La macro se importa a Excel de escritorio y crea borradores mediante la automatización COM de **Outlook clásico en Windows**; no está diseñada para el nuevo Outlook. Revise los destinatarios y adjuntos manualmente antes de enviar. No envía correos automáticamente.

## Seguridad
GitHub Pages es público: nunca suba participantes, correos, PDFs firmados ni credenciales al repositorio. La aplicación procesa los datos localmente en el navegador. Usar solo dentro de políticas institucionales. Los archivos cargados no se guardan entre sesiones. El uso de bibliotecas CDN requiere conexión a Internet.

## Actualización: asociación automática y QR sin CDN
- Cargar PDF múltiples o un ZIP de PDF firmados.
- Coincidencias automáticas por nombre de archivo, sin menús por persona.
- El QR se genera con `qr-local.js` incluido; ya no necesita la biblioteca QRCode remota.
- Excel A: nombre; B: enlace HTTPS de Drive.
- Se reportan excepciones sin obligar a seleccionar archivos manualmente.
- Los otros recursos externos XLSX, JSZip, PDFLib sí siguen dependiendo de Internet.
- Insertar QR después de la firma puede invalidar la firma digital: conservar originales.
