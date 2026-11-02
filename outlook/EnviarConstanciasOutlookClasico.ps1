# Gestor Constancias MIMP - envio directo desde Outlook clasico Windows.
# Envia solo tras confirmacion explicita. No elude politicas de ejecucion.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csvFile = Join-Path $root 'ENVIO_OUTLOOK_LOCAL.csv'
$folder = Join-Path $root 'DIPLOMAS_FINALES_QR'
if (!(Test-Path -LiteralPath $csvFile)) { throw 'Falta ENVIO_OUTLOOK_LOCAL.csv. Extraiga TODO el ZIP antes de ejecutar.' }
if (!(Test-Path -LiteralPath $folder)) { throw 'Falta DIPLOMAS_FINALES_QR. Extraiga TODO el ZIP antes de ejecutar.' }
$items = @(Import-Csv -LiteralPath $csvFile -Delimiter ';' -Encoding UTF8)
if ($items.Count -eq 0) { throw 'No hay participantes en el control de envios.' }
Write-Host ('Registros: '+$items.Count)
Write-Host 'ATENCION: esta herramienta ENVIARA correos reales tras su confirmacion.'
$requested = Read-Host 'Cuenta emisora visible en Outlook (ejemplo cendoc@mimp.gob.pe)'
if ([string]::IsNullOrWhiteSpace($requested)) { throw 'Cuenta emisora obligatoria: se cancela para evitar elegir la cuenta incorrecta.' }
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch { throw 'Abra Outlook CLASICO en Windows antes de ejecutar. El nuevo Outlook no admite COM.' }
$account = $null
foreach ($a in $ol.Session.Accounts) { if ([string]$a.SmtpAddress -ieq $requested.Trim()) { $account = $a; break } }
if ($null -eq $account) { throw 'Cuenta no encontrada entre las cuentas configuradas en Outlook clasico. No se envio ningun mensaje. Si es buzon compartido, consulte a OGTI.' }
$valid=@(); $failed=@()
foreach($item in $items) {
 $name = [string]$item.Nombre
 $email = [string]$item.Correo
 $file = [string]$item.PDF
 if ($email -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { $failed += "$name | correo invalido"; continue }
 if ([IO.Path]::GetFileName($file) -ne $file -or !$file.ToLowerInvariant().EndsWith('.pdf')) { $failed += "$name | nombre PDF invalido"; continue }
 $full=Join-Path $folder $file
 if (!(Test-Path -LiteralPath $full -PathType Leaf)) { $failed += "$name | PDF no encontrado: $file"; continue }
 $valid += [pscustomobject]@{Nombre=$name;Correo=$email;PDF=$file;Ruta=$full;Enlace=[string]$item.Enlace}
}
Write-Host ('Validos: '+$valid.Count+'; requieren revision: '+$failed.Count)
if ($failed.Count) { $failed | Out-Host }
if ($valid.Count -eq 0) { throw 'No hay correos validos para enviar.' }
Write-Host 'Compruebe los destinatarios, adjuntos y cuenta remitente antes de confirmar.'
$answer=Read-Host ('Escriba ENVIAR para enviar AHORA '+$valid.Count+' correos reales desde '+$requested)
if ($answer -cne 'ENVIAR') { Write-Host 'Cancelado.'; exit }
$created=0; $log=New-Object System.Collections.Generic.List[object]
foreach ($item in $items) { if ([string]$item.Correo -notmatch '^[^\s@]+@[^\s@]+\.[^\s@]+$') { Write-Warning ('Correo invalido: ' + $item.Correo) } }
foreach($item in $valid) {
 try {
  $mail=$ol.CreateItem(0)
  $mail.SendUsingAccount=$account
  $mail.To=$item.Correo
  $mail.Subject='Constancia de participacion - MIMP'
  $mail.Body="Estimado(a) $($item.Nombre):`r`n`r`nReciba un cordial saludo. Por medio del presente remitimos su constancia de participacion con codigo QR.`r`n`r`nEnlace de consulta: $($item.Enlace)`r`n`r`nAtentamente,`r`nOficina de Atencion a la Ciudadania y Gestion Documental - MIMP"
  $null=$mail.Attachments.Add($item.Ruta)
  $mail.Send()
  $created++
  $log.Add([pscustomobject]@{Nombre=$item.Nombre;Correo=$item.Correo;Estado='ACEPTADO_POR_OUTLOOK'})
 } catch { $log.Add([pscustomobject]@{Nombre=$item.Nombre;Correo=$item.Correo;Estado=('ERROR: '+$_.Exception.Message)}) }
}
$log | Export-Csv -LiteralPath (Join-Path $root ('REPORTE_ENVIO_'+(Get-Date -Format 'yyyyMMdd_HHmmss')+'.csv')) -NoTypeInformation -Delimiter ';' -Encoding UTF8
Write-Host ('Solicitudes aceptadas por Outlook: '+$created+'. Revise Elementos enviados y Bandeja de salida; la aceptacion no garantiza entrega al destinatario.')
Write-Host 'Se ha creado un REPORTE_ENVIO con fecha y hora. NO ejecute otra vez todo el paquete sin comprobar el reporte: podria duplicar envios.'
