Attribute VB_Name = "ConstanciasMIMP"
Option Explicit

' Ejecutar desde Excel de escritorio en Windows con Outlook clasico instalado.
' Crea BORRADORES, nunca envia mensajes automaticamente.
Public Sub PrepararBorradoresConstancias()
    On Error GoTo ErrorGeneral
    Dim archivo As Variant, carpeta As String
    Dim dlg As FileDialog, libro As Workbook, hoja As Worksheet
    Dim app As Object, mail As Object
    Dim ultima As Long, fila As Long, contador As Long, omitidos As Long
    Dim nombre As String, destinatario As String, enlace As String
    Dim pdf As String, qr As String, rutaPDF As String, rutaQR As String
    Dim cuenta As Object, cuentaDeseada As String

    archivo = Application.GetOpenFilename("Excel (*.xlsx),*.xlsx", , "Seleccione CONTROL_ENVIO_OUTLOOK.xlsx")
    If VarType(archivo) = vbBoolean Then Exit Sub
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    dlg.Title = "Seleccione la carpeta raiz del paquete descomprimido"
    If dlg.Show <> -1 Then Exit Sub
    carpeta = dlg.SelectedItems(1)
    Set libro = Workbooks.Open(CStr(archivo), ReadOnly:=False)
    Set hoja = libro.Worksheets("Envios")
    Set app = GetObject(, "Outlook.Application")
    ultima = hoja.Cells(hoja.Rows.Count, "A").End(xlUp).Row

    cuentaDeseada = InputBox("Cuenta SMTP emisora (opcional)." & vbCrLf & _
         "Deje vacio para utilizar la cuenta predeterminada de Outlook.", "Cuenta de envio")

    For fila = 2 To ultima
        nombre = Trim(CStr(hoja.Cells(fila, 1).Value))
        destinatario = Trim(CStr(hoja.Cells(fila, 2).Value))
        enlace = Trim(CStr(hoja.Cells(fila, 3).Value))
        pdf = Trim(CStr(hoja.Cells(fila, 4).Value))
        qr = Trim(CStr(hoja.Cells(fila, 5).Value))
        rutaPDF = carpeta & "\PDF_FIRMADOS\" & pdf
        rutaQR = carpeta & "\QR_COMPLEMENTARIOS\" & qr
        If Len(nombre) = 0 Or InStr(destinatario, "@") = 0 Or _
           Dir(rutaPDF) = "" Or Dir(rutaQR) = "" Then
            hoja.Cells(fila, 6).Value = "REVISAR: datos o archivos faltantes"
            omitidos = omitidos + 1
        Else
            Set mail = app.CreateItem(0)
            mail.To = destinatario
            mail.Subject = "Constancia de participacion - MIMP"
            mail.Body = "Estimado(a) " & nombre & ":" & vbCrLf & vbCrLf & _
               "Por medio del presente, compartimos su constancia de participacion " & _
               "firmada digitalmente, junto con una hoja complementaria que contiene " & _
               "el codigo QR de acceso al documento." & vbCrLf & vbCrLf & _
               "Enlace: " & enlace & vbCrLf & vbCrLf & _
               "Atentamente," & vbCrLf & _
               "Oficina de Atencion a la Ciudadania y Gestion Documental - MIMP"
            mail.Attachments.Add rutaPDF
            mail.Attachments.Add rutaQR
            If Len(cuentaDeseada) > 0 Then
                For Each cuenta In app.Session.Accounts
                    If LCase(Trim(cuenta.SmtpAddress)) = LCase(cuentaDeseada) Then
                        Set mail.SendUsingAccount = cuenta
                        Exit For
                    End If
                Next cuenta
            End If
            mail.Save
            hoja.Cells(fila, 6).Value = "BORRADOR CREADO"
            contador = contador + 1
        End If
        Set mail = Nothing
    Next fila
    libro.Save
    MsgBox contador & " borradores creados en Outlook." & vbCrLf & _
           omitidos & " filas requieren revision." & vbCrLf & _
           "Revise los destinatarios y adjuntos antes de enviar.", vbInformation
    Exit Sub
ErrorGeneral:
    MsgBox "No se pudo completar el proceso: " & Err.Description & vbCrLf & _
           "Requiere Excel y Outlook clasico de escritorio para Windows.", vbExclamation
End Sub
