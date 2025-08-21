' Autor Fernando Dorantes Nieto
Sub EnviarInvitacionAbiertaDesdeExcel()

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object
    Dim rng As Object
    Dim cell As Object
    Dim dlg As FileDialog
    Dim excelWasCreatedByMe As Boolean
    Const xlUp = -4162 ' Constante de Excel para End(xlUp)
    
    ' Obtiene la aplicación de Outlook
    Dim olApp As Outlook.Application
    Set olApp = Outlook.Application
    
    ' Obtiene la invitación de calendario actualmente abierta
    Dim olApt As Outlook.AppointmentItem
    
    On Error Resume Next ' Evita errores si no hay una invitación abierta
    Set olApt = olApp.ActiveInspector.CurrentItem
    On Error GoTo 0
    
    ' --- Verifica si el usuario tiene una invitación de calendario abierta ---
    If olApt Is Nothing Then
        MsgBox "Por favor, crea y abre una invitación de calendario antes de ejecutar esta macro.", vbExclamation
        Exit Sub
    End If
    If olApt.Class <> olAppointment Then
        MsgBox "El elemento activo no es una invitación de calendario.", vbExclamation
        Exit Sub
    End If
    
    ' --- Estrategia Robusta: Crear una instancia de Excel dedicada y aislada ---
    ' En lugar de arriesgarnos a conectar con un proceso "colgado" (GetObject),
    ' creamos una instancia nueva y limpia que controlaremos de principio a fin.
    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    excelWasCreatedByMe = (Err.Number = 0) ' Será True si se creó correctamente
    On Error GoTo ErrorHandler ' Reactivar el manejo de errores normal
    
    If Not excelWasCreatedByMe Then
        MsgBox "Error Crítico al crear la instancia de Excel." & vbCrLf & vbCrLf & _
               "Causa probable: Permisos de DCOM, Antivirus, o una instalación de Office dañada." & vbCrLf & _
               "Intente ejecutar una 'Reparación Rápida' de Office desde el Panel de Control.", vbCritical
        Exit Sub
    End If
    
    xlApp.Visible = True ' Hacemos Excel visible para que el usuario vea qué pasa
    DoEvents ' Permite que la interfaz de usuario se actualice
    
    ' --- 1. SELECCIONAR EL ARCHIVO DE EXCEL ---
    Set dlg = xlApp.FileDialog(msoFileDialogFilePicker)

    With dlg
        .Title = "Selecciona el archivo de Excel con los correos"
        .Filters.Clear
        .Filters.Add "Archivos de Excel", "*.xlsx;*.xlsm;*.xls"
        .AllowMultiSelect = False
        
        If .Show = -1 Then
            Dim rutaArchivoExcel As String
            rutaArchivoExcel = .SelectedItems(1)
        Else
            MsgBox "No se seleccionó ningún archivo. Operación cancelada.", vbExclamation
            Exit Sub
        End If
    End With
    
    ' --- 2. PEDIR EL NOMBRE DE LA HOJA Y LA COLUMNA ---
    Dim nombreHoja As String
    nombreHoja = InputBox("Ingresa el nombre de la hoja de cálculo (ej. Hoja1):", "Nombre de la Hoja")
    If Trim(nombreHoja) = "" Then
        MsgBox "No se ingresó el nombre de la hoja. Operación cancelada.", vbExclamation
        Exit Sub
    End If

    Dim nombreColumnaCorreos As String
    nombreColumnaCorreos = InputBox("Ingresa la letra de la columna con los correos (ej. A, B, C):", "Columna de Correos")
    If Trim(nombreColumnaCorreos) = "" Then
        MsgBox "No se ingresó la columna de correos. Operación cancelada.", vbExclamation
        Exit Sub
    End If

    On Error GoTo ErrorHandler
    Set xlWb = xlApp.Workbooks.Open(rutaArchivoExcel)
    Set xlWs = xlWb.Sheets(nombreHoja)
    
    ' Define el rango de correos
    Set rng = xlWs.Range(nombreColumnaCorreos & "2:" & nombreColumnaCorreos & xlWs.Cells(xlWs.Rows.Count, nombreColumnaCorreos).End(xlUp).Row)
    
    ' Itera sobre cada correo y envía la invitación
    Dim olAptCopy As Outlook.AppointmentItem
    For Each cell In rng
        ' Verifica que la celda contenga un correo válido
        If Trim(cell.Value) <> "" And InStr(cell.Value, "@") > 0 Then
            ' Crea una copia de la invitación original para cada destinatario
            Set olAptCopy = olApt.Copy
            
            ' Agrega el nuevo destinatario a la COPIA
            olAptCopy.Recipients.Add Trim(cell.Value)
            olAptCopy.Recipients.ResolveAll
            
            ' Envía la copia de la invitación
            olAptCopy.Send
            
            ' Libera la memoria de la copia
            Set olAptCopy = Nothing
        End If
    Next cell
    
    ' Cierra la invitación original que sirvió de plantilla
    olApt.Close olDiscard
    
    MsgBox "Invitaciones enviadas exitosamente.", vbInformation
    
    GoTo Cleanup ' Salta al final para limpiar

ErrorHandler:    
    If Err.Number <> 0 Then ' Solo muestra el mensaje si hay un error
        MsgBox "Se ha producido un error: " & Err.Description, vbCritical
    End If
    
Cleanup:
    ' Cierra Excel y limpia los objetos
    If Not xlWb Is Nothing Then
        xlWb.Close False ' Cierra sin guardar cambios
    End If
    If Not xlApp Is Nothing Then
        If excelWasCreatedByMe Then
            xlApp.Quit ' Solo cierra Excel si esta macro lo creó
        End If
    End If
    
    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing
    Set olApt = Nothing
    Set olAptCopy = Nothing
End Sub
