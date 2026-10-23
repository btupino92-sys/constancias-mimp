# Gestor Constancias MIMP - borradores para Outlook clasico Windows.
# No envia mensajes. No elude politicas de ejecucion.
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$csvFile = Join-Path $root 'ENVIO_OUTLOOK_LOCAL.csv'
$folder = Join-Path $root 'DIPLOMAS_FINALES_QR'
if (!(Test-Path -LiteralPath $csvFile)) { throw 'Falta ENVIO_OUTLOOK_LOCAL.csv. Extraiga TODO el ZIP antes de ejecutar.' }
if (!(Test-Path -LiteralPath $folder)) { throw 'Falta DIPLOMAS_FINALES_QR. Extraiga TODO el ZIP antes de ejecutar.' }
$items = @(Import-Csv -LiteralPath $csvFile -Delimiter ';' -Encoding UTF8)
if ($items.Count -eq 0) { throw 'No hay participantes en el control de envios.' }
Write-Host ('Registros: '+$items.Count)
Write-Host 'Esta herramienta crea BORRADORES. No envia correos.'
$requested = Read-Host 'Cuenta emisora visible en Outlook (ejemplo cendoc@mimp.gob.pe)'
if ([string]::IsNullOrWhiteSpace($requested)) { throw 'Cuenta emisora obligatoria: se cancela para evitar elegir la cuenta incorrecta.' }
try { $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application') }
catch { throw 'Abra Outlook CLASICO en Windows antes de ejecutar. El nuevo Outlook no admite COM.' }
$account = $null
foreach ($a in $ol.Session.Accounts) { if ([string]$a.SmtpAddress -ieq $requested.Trim()) { $account = $a; break } }
if ($null -eq $account) { throw 'Cuenta no encontrada entre las cuentas configuradas en Outlook clasico. No se crearon borradores. Si es buzon compartido, consulte a OGTI.' }
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
if ($valid.Count -eq 0) { throw 'No hay correos validos para crear borradores.' }
$answer=Read-Host ('Escriba CREAR para preparar '+$valid.Count+' borradores desde '+$requested+' (no se envian)')
if ($answer -cne 'CREAR') { Write-Host 'Cancelado.'; exit }
$created=0; $log=New-Object System.Collections.Generic.List[object]
foreach($item in $valid) {
 try {
  $mail=$ol.CreateItem(0)
  $mail.SendUsingAccount=$account
  $mail.To=$item.Correo
  $mail.Subject='Constancia de participacion - MIMP'
  $mail.Body="Estimado(a) $($item.Nombre):`r`n`r`nReciba un cordial saludo. Por medio del presente remitimos su constancia de participacion con codigo QR.`r`n`r`nEnlace de consulta: $($item.Enlace)`r`n`r`nAtentamente,`r`nOficina de Atencion a la Ciudadania y Gestion Documental - MIMP"
  $null=$mail.Attachments.Add($item.Ruta)
  $mail.Save()
  $created++
  $log.Add([pscustomobject]@{Nombre=$item.Nombre;Correo=$item.Correo;Estado='BORRADOR CREADO'})
 } catch { $log.Add([pscustomobject]@{Nombre=$item.Nombre;Correo=$item.Correo;Estado=('ERROR: '+$_.Exception.Message)}) }
}
$log | Export-Csv -LiteralPath (Join-Path $root 'REPORTE_BORRADORES.csv') -NoTypeInformation -Delimiter ';' -Encoding UTF8
Write-Host ('Borradores creados: '+$created+'. Revise la carpeta Borradores, destinatarios, cuenta y adjuntos antes de enviar.')
Write-Host 'Se ha creado REPORTE_BORRADORES.csv.'
