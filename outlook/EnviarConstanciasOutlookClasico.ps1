# Gestor Constancias MIMP - envio directo desde Outlook clasico Windows.
# Envia solo despues de confirmacion expresa. No elude politicas de ejecucion.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csvFile = Join-Path $root 'ENVIO_OUTLOOK_LOCAL.csv'
$folder = Join-Path $root 'DIPLOMAS_FINALES_QR'
if (!(Test-Path -LiteralPath $csvFile)) { throw 'Falta ENVIO_OUTLOOK_LOCAL.csv. Extraiga TODO el ZIP antes de ejecutar.' }
if (!(Test-Path -LiteralPath $folder)) { throw 'Falta DIPLOMAS_FINALES_QR. Extraiga TODO el ZIP antes de ejecutar.' }
$items = @(Import-Csv -LiteralPath $csvFile -Delimiter ';' -Encoding UTF8)
if ($items.Count -eq 0) { throw 'No hay participantes en el control de envios.' }
Write-Host ('Registros: '+$items.Count)
Write-Host 'Esta herramienta ENVIA CORREOS reales al confirmar. Revise todos los datos.'
$requested = Read-Host 'Cuenta emisora visible en Outlook (ejemplo cendoc@mimp.gob.pe)'
if ([string]::IsNullOrWhiteSpace($requested)) { throw 'Cuenta emisora obligatoria: se cancela para evitar elegir la cuenta incorrecta.' }
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch { throw 'Abra Outlook CLASICO en Windows antes de ejecutar. El nuevo Outlook no admite COM.' }
$account = $null
foreach ($a in $ol.Session.Accounts) { if ([string]$a.SmtpAddress -ieq $requested.Trim()) { $account = $a; break } }
if ($null -eq $account) { throw 'Cuenta no encontrada entre las cuentas configuradas en Outlook clasico. No se enviaron correos. Si es buzon compartido, consulte a OGTI.' }
$valid=@(); $failed=@()
foreach($item in $items) {
 $name = [string]$item.Nombre
 $email = [string]$item.Correo
 $file = [string]$item.PDF
 if ($email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { $failed += "$name | correo invalido"; continue }
 if ([IO.Path]::GetFileName($file) -ne $file -or !$file.ToLowerInvariant().EndsWith('.pdf')) { $failed += "$name | nombre PDF invalido"; continue }
 $full=Join-Path $folder $file
 if (!(Test-Path -LiteralPath $full -PathType Leaf)) { $failed += "$name | PDF no encontrado: $file"; continue }
 $valid += [pscustomobject]@{Nombre=$name;Correo=$email;PDF=$file;Ruta=$full;Enlace=[string]$item.Enlace;Actividad=[string]$item.Actividad}
}
Write-Host ('Validos: '+$valid.Count+'; requieren revision: '+$failed.Count)
if ($failed.Count) { $failed | Out-Host }
if ($valid.Count -eq 0) { throw 'No hay correos validos para enviar.' }
$answer=Read-Host ('Escriba ENVIAR para enviar '+$valid.Count+' correos reales desde '+$requested)
if ($answer -cne 'ENVIAR') { Write-Host 'Cancelado.'; exit }
function Get-ActividadParaCorreo([string]$texto) {
 $texto=($texto -replace '\s+', ' ').Trim()
 if ([string]::IsNullOrWhiteSpace($texto)) { throw 'Falta la actividad en el CSV. Genere un paquete nuevo desde el gestor actualizado.' }
 # El dato procede del campo de actividad de la constancia; se elimina solo la formula introductoria.
 $texto=$texto -replace '(?i)^por\s+su\s+participaci[oó]n\s+en\s+(?:el|la|los|las)\s+', ''
 $texto=$texto -replace '(?i)^por\s+haber\s+participado\s+en\s+(?:el|la|los|las)\s+', ''
 # La fecha es un dato distinto de la denominacion del evento.
 $texto=$texto -replace '(?i),?\s*(?:desarrollad[oa]s?|realizad[oa]s?|llevad[oa]s?\s+a\s+cabo)\s+(?:el\s+)?(?:d[ií]a\s+)?\d{1,2}\s+de\s+.+$', ''
 return $texto.Trim().TrimEnd('.',',',' ')
}
$created=0; $log=New-Object System.Collections.Generic.List[object]
foreach($item in $valid) {
 try {
  $mail=$ol.CreateItem(0)
  $mail.SendUsingAccount=$account
  $mail.To=$item.Correo
  $mail.Subject='Constancia de participacion - Agenda Cultural - CENDOC - MIMP'
  $actividad=Get-ActividadParaCorreo $item.Actividad
  $mail.Body="Estimado(a) $($item.Nombre):`r`n`r`nReciba un cordial saludo.`r`n`r`nSe remite adjunta su constancia de participacion correspondiente a la actividad $actividad, realizada en el marco de la Agenda Cultural del Centro de Documentacion (CENDOC) del Ministerio de la Mujer y Poblaciones Vulnerables.`r`n`r`nEnlace de consulta de la constancia: $($item.Enlace)`r`n`r`nAtentamente,`r`nCentro de Documentacion (CENDOC)`r`nOficina de Atencion a la Ciudadania y Gestion Documental`r`nMinisterio de la Mujer y Poblaciones Vulnerables"
  $null=$mail.Attachments.Add($item.Ruta)
  $mail.Send()
  $created++
  $log.Add([pscustomobject]@{Nombre=$item.Nombre;Correo=$item.Correo;Estado='ACEPTADO_POR_OUTLOOK'})
 } catch { Write-Host ('ERROR para '+$item.Correo+': '+$_.Exception.Message) -ForegroundColor Red; $log.Add([pscustomobject]@{Nombre=$item.Nombre;Correo=$item.Correo;Estado=('ERROR: '+$_.Exception.Message)}) }
}
$log | Export-Csv -LiteralPath (Join-Path $root 'REPORTE_ENVIO.csv') -NoTypeInformation -Delimiter ';' -Encoding UTF8
Write-Host ('Envios aceptados por Outlook: '+$created+'. Revise Elementos enviados y Bandeja de salida; la aceptacion no garantiza entrega.')
Write-Host 'Se ha creado REPORTE_ENVIO.csv. Evite repetir el proceso para no duplicar envios.'
