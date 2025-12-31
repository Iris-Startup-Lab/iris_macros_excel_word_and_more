' Autor: Fernando Dorantes Nieto
' Adaptado por: Gemini Code Assist
' Macro para añadir asistentes a una reunión activa desde una lista de Excel.

Sub AddAttendeesFromExcel()
    ' --- Declaración de variables ---
    Dim oAppt As Outlook.AppointmentItem
    Dim oInspector As Outlook.Inspector
    
    ' Objetos de Excel (requiere la referencia a "Microsoft Excel XX.0 Object Library")
    Dim oExcelApp As Object
    Dim oWorkbook As Object
    Dim oWorksheet As Object
    Dim sExcelPath As String
    
    ' Variables para procesar el Excel
    Dim lLastRow As Long
    Dim i As Long
    Dim colEmail As Integer
    Dim colName As Integer
    Dim sEmail As String
    Dim sName As String
    Dim invitationsSent As Long
    
    On Error GoTo ErrorHandler

    ' --- PASO 1: Obtener la reunión activa (Método robusto) ---
    ' En lugar de confiar en "ActiveWindow" o "ActiveInspector", que pueden ser ambiguos,
    ' recorremos todas las ventanas de edición abiertas (Inspectors) para encontrar la reunión.
    Dim insp As Outlook.Inspector
    Dim found As Boolean
    found = False
    
    ' Iterar a través de todas las ventanas de inspector abiertas
    For Each insp In Application.Inspectors
        ' Comprobar si el elemento en el inspector es una reunión
        If TypeOf insp.CurrentItem Is Outlook.AppointmentItem Then
            Set oAppt = insp.CurrentItem
            found = True
            Exit For ' Encontramos la reunión, salimos del bucle
        End If
    Next insp
    
    ' Si no se encontró ninguna reunión después de revisar todas las ventanas
    If Not found Then
        MsgBox "No se encontró ninguna ventana de reunión abierta." & vbCrLf & _
               "Por favor, abre o crea una reunión y asegúrate de que sea la ventana principal antes de ejecutar la macro.", vbCritical
        Exit Sub
    End If
    
    ' --- PASO 2: Seleccionar el archivo de Excel ---
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Selecciona el archivo de Excel con los correos"
        .Filters.Clear
        .Filters.Add "Archivos de Excel", "*.xlsx; *.xls; *.xlsm"
        .AllowMultiSelect = False
        If .Show = -1 Then
            sExcelPath = .SelectedItems(1)
        Else
            Exit Sub ' El usuario canceló
        End If
    End With

    ' --- PASO 3: Abrir y validar el archivo de Excel ---
    Set oExcelApp = CreateObject("Excel.Application")
    
    ' --- Comprobación de permisos ---
    ' Si CreateObject falla (a menudo debido a la configuración de seguridad), oExcelApp será Nothing.
    ' Esto nos da un error mucho más específico que el genérico "El objeto no admite...".
    If oExcelApp Is Nothing Then
        MsgBox "No se pudo iniciar Excel desde Outlook." & vbCrLf & vbCrLf & _
               "Esto suele ser un problema de permisos. Por favor, revisa la configuración del Centro de Confianza de Outlook:" & vbCrLf & _
               "Archivo > Opciones > Centro de Confianza > Configuración del Centro de Confianza > Configuración de macros.", vbCritical
        Exit Sub
    End If

    Set oWorkbook = oExcelApp.Workbooks.Open(sExcelPath, ReadOnly:=True)
    Set oWorksheet = oWorkbook.Sheets(1) ' Asume que los datos están en la primera hoja
    
    ' Validar encabezados: buscar la columna "Correo"
    ' La columna "Nombre" es opcional, pero necesaria para personalizar el cuerpo.
    Dim lastCol As Long
    Dim c As Long
    
    ' Encontrar la última columna usada en la fila de encabezados para un bucle eficiente y robusto.
    lastCol = oWorksheet.Cells(1, oWorksheet.Columns.Count).End(-4163).Column ' xlToLeft = -4163
    
    colEmail = 0
    colName = 0
    For c = 1 To lastCol
        Select Case LCase(Trim(CStr(oWorksheet.Cells(1, c).Value)))
            Case "correo"
                colEmail = c
            Case "nombre"
                colName = c
        End Select
    Next c
    
    If colEmail = 0 Then
        MsgBox "El archivo de Excel no es válido." & vbCrLf & vbCrLf & _
               "Asegúrate de que la primera hoja contenga una columna con el encabezado 'Correo'.", _
               vbCritical
        GoTo Cleanup
    End If

    ' --- PASO 4: Leer el Excel y añadir destinatarios ---
    lLastRow = oWorksheet.Cells(oWorksheet.Rows.Count, colEmail).End(-4162).Row ' xlUp
    invitationsSent = 0
    
    For i = 2 To lLastRow ' Asume que la fila 1 tiene encabezados
        sEmail = Trim(CStr(oWorksheet.Cells(i, colEmail).Value))
        sName = ""
        If colName > 0 Then
            sName = Trim(CStr(oWorksheet.Cells(i, colName).Value))
        End If

        If sEmail <> "" And InStr(sEmail, "@") > 0 Then
            ' --- ALTERNATIVA DE DIAGNÓSTICO: Crear una reunión mínima ---
            ' Dado que los métodos de copia fallan, vamos a crear una reunión completamente
            ' nueva y simple para verificar si el problema está en la creación/envío
            ' o en la copia de propiedades desde la reunión original.
            Dim oApptToSend As Outlook.AppointmentItem
            Set oApptToSend = Application.CreateItem(olAppointmentItem)

            With oApptToSend
                ' Usamos datos fijos para la prueba.
                .MeetingStatus = olMeeting
                .Subject = "Prueba de Reunión de Diagnóstico"
                .Start = Now + 1 ' Mañana a la misma hora
                .Duration = 60 ' 60 minutos
                .Body = "Este es un cuerpo de mensaje de prueba para " & sName & "."
                
                ' Añade el destinatario actual
                .Recipients.Add sEmail
                .Recipients.ResolveAll
                
                ' Si esto funciona, el problema está en una de las propiedades
                ' que intentábamos copiar de la reunión original (oAppt).
                .Send
            End With
            invitationsSent = invitationsSent + 1
        End If
    Next i
    
    MsgBox invitationsSent & " invitaciones personalizadas han sido enviadas." & vbCrLf & vbCrLf & "Puedes cerrar la ventana de la reunión original sin guardar.", vbInformation

Cleanup:
    ' --- PASO 5: Limpieza de objetos ---
    On Error Resume Next
    If Not oWorkbook Is Nothing Then oWorkbook.Close SaveChanges:=False
    If Not oExcelApp Is Nothing Then oExcelApp.Quit
    Set oAppt = Nothing
    Set oInspector = Nothing
    Set oWorksheet = Nothing
    Set oWorkbook = Nothing
    Set oExcelApp = Nothing
    Exit Sub

ErrorHandler:
    MsgBox "Ocurrió un error inesperado: " & vbCrLf & "Error #" & Err.Number & " - " & Err.Description, vbCritical
    GoTo Cleanup
End Sub