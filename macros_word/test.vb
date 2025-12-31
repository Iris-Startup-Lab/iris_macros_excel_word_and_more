' Autor Fernando Dorantes Nieto
Dim rutaICS As String
Dim asuntoCorreo As String
Dim rutasImagenes() As String

Sub PrepararEnvioConICS()
    Dim dlgICS As FileDialog
    Dim dlgImagenes As FileDialog
    Dim i As Integer

    ' Seleccionar archivo .ics
    Set dlgICS = Application.FileDialog(msoFileDialogFilePicker)
    With dlgICS
        .Title = "Selecciona el archivo .ics del evento"
        .Filters.Clear
        .Filters.Add "Archivos de calendario", "*.ics"
        .AllowMultiSelect = False

        If .Show = -1 Then
            rutaICS = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ningún archivo .ics.", vbExclamation
            Exit Sub
        End If
    End With

    ' Preguntar por el asunto del correo
    asuntoCorreo = InputBox("Escribe el asunto del correo electrónico:", "Asunto del mensaje")
    If asuntoCorreo = "" Then
        MsgBox "No se ingresó ningún asunto.", vbExclamation
        Exit Sub
    End If

    ' Seleccionar imágenes opcionales
    Set dlgImagenes = Application.FileDialog(msoFileDialogFilePicker)
    With dlgImagenes
        .Title = "Selecciona una o más imágenes para adjuntar (opcional)"
        .Filters.Clear
        .Filters.Add "Imágenes", "*.jpg; *.jpeg; *.png; *.gif"
        .AllowMultiSelect = True

        If .Show = -1 Then
            ReDim rutasImagenes(.SelectedItems.Count - 1)
            For i = 1 To .SelectedItems.Count
                rutasImagenes(i - 1) = .SelectedItems(i)
            Next i
        Else
            ' No se seleccionaron imágenes, continuar sin error
            ReDim rutasImagenes(0)
            rutasImagenes(0) = ""
        End If
    End With

    MsgBox "Archivo .ics, asunto y adjuntos listos. Ahora puedes ejecutar el envío.", vbInformation
End Sub

Sub EnviarCorreosconCalendario()
    Dim i As Long, j As Integer
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim doc As Document
    Dim campoCorreo As String
    Dim oAccount As Object ' Para la cuenta de Outlook a usar
    Dim oAccounts As Object ' Colección de cuentas de Outlook
    Dim strAccounts As String ' Lista de cuentas para mostrar al usuario
    Dim selectedAccountEmail As String ' Email de la cuenta seleccionada
    Dim accountFound As Boolean
    Dim cuerpoMensaje As String

    If rutaICS = "" Or asuntoCorreo = "" Then
        MsgBox "Primero debes ejecutar 'PrepararEnvioConICS'.", vbExclamation
        Exit Sub
    End If

    Set OutlookApp = CreateObject("Outlook.Application")

    ' Obtener las cuentas de Outlook configuradas
    Set oAccounts = OutlookApp.Session.Accounts
    If oAccounts.Count = 0 Then
        MsgBox "No se encontraron cuentas de Outlook configuradas.", vbCritical
        Exit Sub
    End If

    ' Si hay más de una cuenta, preguntar al usuario cuál usar
    If oAccounts.Count > 1 Then
        ' Construir la lista de cuentas para mostrar en el InputBox
        For Each oAccount In oAccounts
            strAccounts = strAccounts & oAccount.SmtpAddress & vbCrLf
        Next

        selectedAccountEmail = InputBox("Por favor, escribe o copia la dirección de correo de la cuenta que deseas usar para enviar:" & vbCrLf & vbCrLf & strAccounts, "Seleccionar Cuenta de Envío")

        If selectedAccountEmail = "" Then
            MsgBox "No se seleccionó ninguna cuenta. Operación cancelada.", vbExclamation
            Exit Sub
        End If

        ' Encontrar el objeto de la cuenta seleccionada
        For Each oAccount In oAccounts
            If LCase(oAccount.SmtpAddress) = LCase(selectedAccountEmail) Then
                accountFound = True
                Exit For
            End If
        Next

        If Not accountFound Then
            MsgBox "La cuenta de correo '" & selectedAccountEmail & "' no fue encontrada. Por favor, verifica la dirección e inténtalo de nuevo.", vbCritical
            Exit Sub
        End If
    Else
        ' Si solo hay una cuenta, usarla por defecto
        Set oAccount = oAccounts(1)
    End If

    Set doc = ActiveDocument

    With doc.MailMerge
        If .MainDocumentType <> wdEMail Then
            MsgBox "Este documento no está configurado para enviar correos electrónicos.", vbExclamation
            Exit Sub
        End If

        For i = 1 To .DataSource.RecordCount
            .DataSource.ActiveRecord = i
            campoCorreo = .DataSource.DataFields("Correo").Value
            cuerpoMensaje = doc.Content.Text

            Set OutlookMail = OutlookApp.CreateItem(0)
            With OutlookMail
                ' Especificar la cuenta desde la que se enviará el correo
                Set .SendUsingAccount = oAccount

                .To = campoCorreo
                .Subject = asuntoCorreo
                .Body = cuerpoMensaje
                .Attachments.Add rutaICS

                ' Adjuntar imágenes si hay
                If rutasImagenes(0) <> "" Then
                    For j = LBound(rutasImagenes) To UBound(rutasImagenes)
                        .Attachments.Add rutasImagenes(j)
                    Next j
                End If

                .Send
            End With
        Next i
    End With

    MsgBox "Correos enviados exitosamente con el archivo .ics y adjuntos.", vbInformation
End Sub
