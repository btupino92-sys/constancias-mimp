# Gestor de Constancias MIMP — versión 1

Aplicación estática HTML/CSS/JS para GitHub Pages. Se utiliza como fondo una versión sin campos variables del modelo DOCM que proporcionó el usuario. La plantilla conserva logotipos, marco, nombre y cargo del director. Genera PDF individuales con nombres completos y descarga un ZIP; permite importar/exportar participantes en Excel y registrar enlaces de Drive; genera QR como PNG independientes y prepara correos mediante `mailto:`.

## Instalación GitHub Pages

1. Crea un repositorio privado o público conforme a las normas internas de tu entidad; GitHub Pages de repositorios públicos expone los archivos de la web públicamente.
2. Sube **index.html**, **styles.css**, **app.js** y la carpeta **assets/** conservando su estructura.
3. GitHub → Settings → Pages → Deploy from branch → `main` → `/(root)` → Save.
4. Abre la URL de GitHub Pages del repositorio, no el archivo `index.html` local (la plantilla PDF se carga mediante `fetch`).

## Uso

1. Revisa actividad y fecha de emisión.
2. Importa Excel con `Nombre completo`, `Correo` y opcional `Enlace`; también puedes agregar participantes manualmente.
3. Previsualiza el primer PDF y descarga `CONSTANCIAS_MIMP_SIN_FIRMA.zip`.
4. Firma externamente los PDF según el procedimiento institucional y luego súbelos a Google Drive.
5. Asigna los enlaces públicos o de acceso restringido según autorización institucional, pega enlaces en el listado o importa un Excel con la columna `Enlace`.
6. Exporta el control de Excel; descarga los QR en PNG (independientes del PDF firmado); usa «Preparar correo» para abrir un borrador individual.

## Limitaciones y seguridad

- No autentica usuarios, ni integra la firma digital, ni sube archivos a Google Drive, ni envía correos de forma autónoma. Estas tareas requieren servicios, credenciales, autorizaciones y controles de seguridad aprobados por la institución.
- Evita incluir listados de nombres, correos o documentos firmados dentro del repositorio: se procesan en memoria en el navegador y no se almacenan automáticamente en GitHub.
- Las bibliotecas JS se cargan de CDN públicos; hace falta conexión a Internet. Un despliegue institucional puede fijar versiones y alojar las dependencias localmente.
- Los QR hacia Drive **no** deben incrustarse después de la firma en los PDF firmados: modificar un PDF firmado puede invalidar su firma. Descarga los códigos QR separadamente o diseña un enlace de verificación estable para incluirlo **antes** de firmar.
- La vista previa depende del lector PDF del navegador. No se realiza validación criptográfica de firma.
- El bloque del director es parte de la plantilla fija. Cambiar director, diseño o logos requiere actualizar la plantilla.
- La importación con nombres duplicados es válida; la exportación agrega sufijos `_2`, `_3` para evitar sobrescrituras. La asociación de enlaces por nombre puede ser ambigua para homónimos: revisar manualmente.
