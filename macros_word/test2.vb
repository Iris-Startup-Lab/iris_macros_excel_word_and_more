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
    Dim htmlMensaje As String
    Dim tempHTMLFile As String
    Dim FSO As Object ' Declara la variable
    Dim ts As Object

    If rutaICS = "" Or asuntoCorreo = "" Then
        MsgBox "Primero debes ejecutar 'PrepararEnvioConICS'.", vbExclamation
        Exit Sub
    End If

    Set OutlookApp = CreateObject("Outlook.Application")
    Set doc = ActiveDocument

    With doc.MailMerge
        If .MainDocumentType <> wdEMail Then
            MsgBox "Este documento no está configurado para enviar correos electrónicos.", vbExclamation
            Exit Sub
        End If

        For i = 1 To .DataSource.RecordCount
            .DataSource.ActiveRecord = i
            campoCorreo = .DataSource.DataFields("Correo").Value
            
            tempHTMLFile = Environ("temp") & "\temp_mail_body.htm"
            
            ' ¡IMPORTANTE! Inicializa el objeto FSO aquí
            Set FSO = CreateObject("Scripting.FileSystemObject")
            
            On Error Resume Next ' Ignora errores si no se puede guardar el archivo
            doc.SaveAs2 tempHTMLFile, wdFormatHTML
            On Error GoTo 0
            
            ' Verifica si el archivo se creó antes de continuar
            If FSO.FileExists(tempHTMLFile) Then
                ' Lee el contenido del archivo HTML
                Set ts = FSO.OpenTextFile(tempHTMLFile, 1)
                htmlMensaje = ts.ReadAll
                ts.Close
                
                ' Libera el objeto FSO antes de salir del bucle
                Set FSO = Nothing
                
                Set OutlookMail = OutlookApp.CreateItem(0)
                With OutlookMail
                    .To = campoCorreo
                    .Subject = asuntoCorreo
                    .HTMLBody = htmlMensaje
                    .Attachments.Add rutaICS

                    If rutasImagenes(0) <> "" Then
                        For j = LBound(rutasImagenes) To UBound(rutasImagenes)
                            .Attachments.Add rutasImagenes(j)
                        Next j
                    End If

                    .Send
                End With
                
                On Error Resume Next
                Kill tempHTMLFile
                On Error GoTo 0
            Else
                MsgBox "No se pudo crear el archivo HTML temporal para el correo. Verifique los permisos.", vbExclamation
            End If
            
        Next i
    End With

    MsgBox "Correos enviados exitosamente con el archivo .ics y adjuntos.", vbInformation
End Sub
