# Envío desde el HTML con Microsoft 365

La página es un cliente web estático. No usa contraseñas ni claves secretas.

## Lo que debe configurar OGTI

1. Registrar una aplicación en Microsoft Entra ID (registro de aplicaciones), con plataforma **Single-page application (SPA)** y URI de redirección exacta a la página publicada de GitHub Pages, por ejemplo `https://usuario.github.io/constancias-mimp/`.
2. Proporcionar **Application (client) ID** y **Directory (tenant) ID**. Son identificadores públicos, no contraseñas.
3. Permitir permisos delegados `User.Read`, `Mail.Send` y `Mail.Send.Shared`, con consentimiento institucional cuando la política lo requiera.
4. Si se envía desde `cendoc@mimp.gob.pe` con inicio de sesión de otra cuenta, configurar permiso de Exchange Online **Send As** o **Send on Behalf** sobre el buzón institucional.
5. Verificar si las políticas del organismo permiten envío de adjuntos por Microsoft Graph desde una página alojada en GitHub Pages.

## Uso

1. Completa pasos 1 a 4 como siempre. El Excel de participantes del paso 2 debe incluir correos válidos.
2. En paso 5 escribe los dos IDs de OGTI; conecta tu cuenta.
3. Pulsa **Preparar lista de envíos**, revisa los destinatarios y desmarca los que no deban recibir correo.
4. Pulsa **Revisar y enviar** y confirma. El envío es real.
5. Los estados ENVIADO significan que Graph aceptó la solicitud (HTTP 202), no que haya evidencia de entrega. Los estados del navegador no se guardan después de cerrar/recargar la página. Para auditoría institucional hace falta una base de datos/log aprobado.

## Límites

- Este cliente usa adjuntos simples, de hasta unos 2.8 MB por PDF.
- Una sesión de navegador no es un sistema fiable de envíos masivos desatendidos; pueden aplicar límites de Microsoft 365.
- La inserción del QR posterior a la firma puede invalidar la firma digital. Se recomienda insertar el QR antes de firmar.
- Nunca publiques participantes reales, credenciales o PDF institucionales en GitHub.
