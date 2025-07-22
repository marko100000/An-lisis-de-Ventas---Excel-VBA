Attribute VB_Name = "Módulo2"
Sub EnviarCorreoConAdjunto()

    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim archivoAdjunto As String
    Dim destinatario As String
    Dim asunto As String
    Dim cuerpo As String

    ' Establecer el archivo adjunto
    archivoAdjunto = ThisWorkbook.FullName ' Esto adjuntará el archivo Excel completo (puedes cambiarlo a cualquier archivo que desees)
    
    ' Definir el destinatario, asunto y cuerpo del mensaje
    destinatario = "MARKO.NAVEDA@UPSJB.EDU.PE" ' Cambia esta dirección por la real
    asunto = "Reporte de Ventas - " & Date ' Asunto del correo
    cuerpo = "Estimado, adjunto el reporte de ventas. Quedo atento a tus comentarios." ' Cuerpo del correo

    ' Crear una instancia de Outlook
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0) ' 0 significa un correo nuevo

    ' Crear el correo con sus propiedades
    With OutlookMail
        .To = destinatario
        .Subject = asunto
        .Body = cuerpo
        .Attachments.Add archivoAdjunto ' Adjuntar el archivo
        .Send ' Enviar el correo
    End With

    ' Limpiar las variables
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing

    MsgBox "Correo enviado exitosamente", vbInformation

End Sub

