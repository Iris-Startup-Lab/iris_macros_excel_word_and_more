' ==============================================================================
' Autor: Fernando Dorantes Nieto & Equipo Iris Startup Lab
' Herramientas: Microsoft Excel + Microsoft Outlook + Microsoft Word (Opcional)
' Descripción: Macro multiherramienta para automatizar el envío de recordatorios
'              y listas de verificación (checklist) a las personas involucradas en
'              un experimento exactamente 1 día antes de su realización.
' ==============================================================================

Option Explicit

' ------------------------------------------------------------------------------
' CONSTANTES DE CONFIGURACIÓN DE LA HOJA
' ------------------------------------------------------------------------------
' Ajusta estos nombres según la estructura de tu archivo de Excel:
Public Const NOMBRE_HOJA_EXPERIMENTOS As String = "Experimentos"
Public Const FILA_INICIAL_DATOS As Long = 2

' Índices de columnas (1 = A, 2 = B, 3 = C, etc.)
Public Const COL_ID As Integer = 1             ' Columna A: ID Experimento
Public Const COL_TITULO As Integer = 2         ' Columna B: Nombre / Título del Experimento
Public Const COL_LIDER As Integer = 3          ' Columna C: Responsable / Líder
Public Const COL_CORREOS As Integer = 4        ' Columna D: Correos involucrados (separados por ;)
Public Const COL_FECHA As Integer = 5          ' Columna E: Fecha del experimento (DD/MM/AAAA)
Public Const COL_HORA As Integer = 6           ' Columna F: Hora / Turno
Public Const COL_UBICACION As Integer = 7      ' Columna G: Laboratorio / Sala / Ubicación
Public Const COL_CHECKLIST As Integer = 8      ' Columna H: Requerimientos / Checklist (saltos de línea)
Public Const COL_ESTADO_ENVIO As Integer = 9   ' Columna I: Estado ("PENDIENTE", "ENVIADO", etc.)
Public Const COL_FECHA_ENVIO As Integer = 10   ' Columna J: Fecha y hora exacta de envío

' ==============================================================================
' PROCEDIMIENTO PRINCIPAL: EnviarRecordatoriosExperimentos
' ==============================================================================
' Recorre la tabla de experimentos, detecta aquellos cuya fecha sea mañana
' (o configurable), verifica que no hayan sido notificados y genera el correo en Outlook.
' ==============================================================================
Sub EnviarRecordatoriosExperimentos()
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim fechaExp As Variant
    Dim diasDiferencia As Long
    Dim correosDestino As String
    Dim tituloExp As String
    Dim responsableExp As String
    Dim horaExp As String
    Dim ubicacionExp As String
    Dim checklistTexto As String
    Dim estadoEnvio As String
    Dim totalEnviados As Long
    Dim totalOmitidos As Long
    
    ' Objetos de Outlook (Late Binding para evitar problemas de referencias)
    Dim olApp As Object
    Dim olMail As Object
    Dim outlookIniciadoPorMacro As Boolean
    
    On Error GoTo ManejoErrores
    
    ' 1. Validar que exista la hoja de experimentos
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(NOMBRE_HOJA_EXPERIMENTOS)
    On Error GoTo ManejoErrores
    
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja '" & NOMBRE_HOJA_EXPERIMENTOS & "' en este libro." & vbCrLf & _
               "Verifica el nombre de la hoja en la constante NOMBRE_HOJA_EXPERIMENTOS.", vbCritical, "Error de Configuración"
        Exit Sub
    End If
    
    ' 2. Encontrar la última fila con datos en la columna B (Título)
    ultimaFila = ws.Cells(ws.Rows.Count, COL_TITULO).End(-4162).Row ' -4162 = xlUp
    
    If ultimaFila < FILA_INICIAL_DATOS Then
        MsgBox "No hay registros de experimentos en la hoja '" & NOMBRE_HOJA_EXPERIMENTOS & "'.", vbInformation, "Sin Datos"
        Exit Sub
    End If
    
    ' 3. Conectar o iniciar Microsoft Outlook
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then
        Set olApp = CreateObject("Outlook.Application")
        outlookIniciadoPorMacro = True
    End If
    On Error GoTo ManejoErrores
    
    If olApp Is Nothing Then
        MsgBox "No fue posible conectar con Microsoft Outlook. Asegúrate de tenerlo instalado y configurado.", vbCritical, "Error de Outlook"
        Exit Sub
    End If
    
    totalEnviados = 0
    totalOmitidos = 0
    
    ' 4. Recorrer cada fila de experimentos
    For i = FILA_INICIAL_DATOS To ultimaFila
        tituloExp = Trim(CStr(ws.Cells(i, COL_TITULO).Value))
        fechaExp = ws.Cells(i, COL_FECHA).Value
        estadoEnvio = UCase(Trim(CStr(ws.Cells(i, COL_ESTADO_ENVIO).Value)))
        correosDestino = Trim(CStr(ws.Cells(i, COL_CORREOS).Value))
        
        ' Solo procesar si hay título, fecha válida y correos
        If tituloExp <> "" And IsDate(fechaExp) And correosDestino <> "" Then
            
            ' Calcular diferencia en días entre la fecha del experimento y HOY
            ' diasDiferencia = 1 significa que el experimento es MAÑANA
            diasDiferencia = DateDiff("d", Date, CDate(fechaExp))
            
            ' Condición: Falta exactamente 1 día y no ha sido marcado como ENVIADO
            If diasDiferencia = 1 And InStr(estadoEnvio, "ENVIADO") = 0 Then
                
                responsableExp = Trim(CStr(ws.Cells(i, COL_LIDER).Value))
                horaExp = Trim(CStr(ws.Cells(i, COL_HORA).Value))
                ubicacionExp = Trim(CStr(ws.Cells(i, COL_UBICACION).Value))
                checklistTexto = Trim(CStr(ws.Cells(i, COL_CHECKLIST).Value))
                
                If horaExp = "" Then horaExp = "Por definir"
                If ubicacionExp = "" Then ubicacionExp = "Laboratorio Principal"
                
                ' Crear correo electrónico en Outlook
                Set olMail = olApp.CreateItem(0) ' 0 = olMailItem
                
                With olMail
                    .To = correosDestino
                    .Subject = "🔔 RECORDATORIO [1 DÍA]: Experimento '" & tituloExp & "' - Checklist y Requerimientos"
                    
                    ' Construir el cuerpo con formato HTML profesional y checklist interactivo
                    .HTMLBody = GenerarPlantillaHTML(tituloExp, responsableExp, CDate(fechaExp), horaExp, ubicacionExp, checklistTexto)
                    
                    ' Descomentar .Display si se prefiere ver antes de enviar, o .Send para envío directo:
                    .Send
                    '.Display
                End With
                
                ' Registrar estatus en el Excel
                ws.Cells(i, COL_ESTADO_ENVIO).Value = "ENVIADO"
                ws.Cells(i, COL_FECHA_ENVIO).Value = Now
                ws.Cells(i, COL_ESTADO_ENVIO).Interior.Color = RGB(198, 239, 206) ' Verde claro
                ws.Cells(i, COL_ESTADO_ENVIO).Font.Color = RGB(0, 97, 0)
                
                totalEnviados = totalEnviados + 1
            Else
                totalOmitidos = totalOmitidos + 1
            End If
        End If
    Next i
    
    ' Resumen al finalizar
    MsgBox "Proceso completado con éxito." & vbCrLf & vbCrLf & _
           "• Recordatorios enviados para mañana: " & totalEnviados & vbCrLf & _
           "• Experimentos omitidos (otras fechas o ya enviados): " & totalOmitidos, _
           vbInformation, "Recordatorios de Experimentos"
           
    Exit Sub

ManejoErrores:
    MsgBox "Ocurrió un error inesperado durante la ejecución: " & vbCrLf & _
           "Número: " & Err.Number & vbCrLf & _
           "Descripción: " & Err.Description, vbCritical, "Error en Macro"
End Sub

' ==============================================================================
' FUNCIÓN AUXILIAR: GenerarPlantillaHTML
' ==============================================================================
' Genera un diseño de correo electrónico moderno, responsivo y visualmente limpio
' con tabla de resumen y lista de verificación con casillas estilizadas.
' ==============================================================================
Private Function GenerarPlantillaHTML(ByVal titulo As String, _
                                     ByVal responsable As String, _
                                     ByVal fecha As Date, _
                                     ByVal hora As String, _
                                     ByVal ubicacion As String, _
                                     ByVal checklist As String) As String
    Dim html As String
    Dim lineasChecklist() As String
    Dim i As Long
    Dim itemChecklist As String
    Dim checklistFormateado As String
    
    ' Procesar el checklist: separar por saltos de línea (vbCrLf, vbLf o vbCr)
    checklist = Replace(checklist, vbCrLf, vbLf)
    checklist = Replace(checklist, vbCr, vbLf)
    lineasChecklist = Split(checklist, vbLf)
    
    checklistFormateado = ""
    For i = LBound(lineasChecklist) To UBound(lineasChecklist)
        itemChecklist = Trim(lineasChecklist(i))
        ' Eliminar guiones o viñetas previas si existen
        If Left(itemChecklist, 2) = "- " Or Left(itemChecklist, 2) = "* " Then
            itemChecklist = Trim(Mid(itemChecklist, 3))
        End If
        
        If itemChecklist <> "" Then
            checklistFormateado = checklistFormateado & _
                "<tr style='border-bottom: 1px solid #E2E8F0;'>" & _
                "  <td style='padding: 10px 12px; width: 32px; vertical-align: top; text-align: center; font-size: 18px; color: #3B82F6;'>&#x2610;</td>" & _
                "  <td style='padding: 10px 12px; vertical-align: middle; color: #1E293B; font-size: 14px;'>" & itemChecklist & "</td>" & _
                "</tr>"
        End If
    Next i
    
    If checklistFormateado = "" Then
        checklistFormateado = "<tr><td colspan='2' style='padding: 12px; color: #64748B; font-style: italic;'>No se especificaron requerimientos previos adicionales.</td></tr>"
    End If
    
    ' Armado de la plantilla HTML con diseño corporativo
    html = "<!DOCTYPE html>" & _
           "<html>" & _
           "<head><meta charset='UTF-8'></head>" & _
           "<body style='margin: 0; padding: 20px; background-color: #F8FAFC; font-family: Segoe UI, Helvetica, Arial, sans-serif; color: #0F172A;'>" & _
           "  <table align='center' border='0' cellpadding='0' cellspacing='0' width='100%' style='max-width: 620px; background-color: #FFFFFF; border-radius: 12px; overflow: hidden; border: 1px solid #E2E8F0; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);'>" & _
           "    <!-- Encabezado con degradado moderno -->" & _
           "    <tr>" & _
           "      <td style='background: linear-gradient(135deg, #1E3A8A 0%, #3B82F6 100%); padding: 28px 30px; text-align: left; color: #FFFFFF;'>" & _
           "        <span style='display: inline-block; background: rgba(255,255,255,0.2); padding: 4px 12px; border-radius: 20px; font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;'>Aviso de Preparación</span>" & _
           "        <h1 style='margin: 12px 0 6px 0; font-size: 22px; font-weight: 700; color: #FFFFFF; line-height: 1.3;'>Recordatorio: Experimento Mañana</h1>" & _
           "        <p style='margin: 0; font-size: 14px; color: #DBEAFE;'>Por favor revisa los requerimientos para estar alineados antes de la prueba.</p>" & _
           "      </td>" & _
           "    </tr>" & _
           "    <!-- Detalles del Experimento -->" & _
           "    <tr>" & _
           "      <td style='padding: 24px 30px;'>" & _
           "        <h2 style='font-size: 17px; margin: 0 0 16px 0; color: #1E293B; border-bottom: 2px solid #F1F5F9; padding-bottom: 8px;'>📋 Ficha del Experimento</h2>" & _
           "        <table border='0' cellpadding='6' cellspacing='0' width='100%' style='font-size: 14px;'>" & _
           "          <tr><td style='width: 140px; color: #64748B; font-weight: 600;'>Experimento:</td><td style='color: #0F172A; font-weight: bold;'>" & titulo & "</td></tr>" & _
           "          <tr><td style='color: #64748B; font-weight: 600;'>Responsable:</td><td style='color: #0F172A;'>" & responsable & "</td></tr>" & _
           "          <tr><td style='color: #64748B; font-weight: 600;'>Fecha programada:</td><td style='color: #2563EB; font-weight: bold;'>" & Format(fecha, "dddd, dd \de mmmm \de yyyy") & "</td></tr>" & _
           "          <tr><td style='color: #64748B; font-weight: 600;'>Horario:</td><td style='color: #0F172A;'>" & hora & "</td></tr>" & _
           "          <tr><td style='color: #64748B; font-weight: 600;'>Ubicación / Sala:</td><td style='color: #0F172A;'>" & ubicacion & "</td></tr>" & _
           "        </table>" & _
           "      </td>" & _
           "    </tr>" & _
           "    <!-- Checklist / Requerimientos Previos -->" & _
           "    <tr>" & _
           "      <td style='padding: 0 30px 24px 30px;'>" & _
           "        <h2 style='font-size: 17px; margin: 0 0 12px 0; color: #1E293B; border-bottom: 2px solid #F1F5F9; padding-bottom: 8px;'>✅ Checklist & Requerimientos Previos</h2>" & _
           "        <p style='font-size: 13px; color: #64748B; margin: 0 0 12px 0;'>Verifica que cada uno de los siguientes puntos esté completado antes de iniciar la sesión:</p>" & _
           "        <table border='0' cellpadding='0' cellspacing='0' width='100%' style='background-color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; overflow: hidden;'>" & _
                      checklistFormateado & _
           "        </table>" & _
           "      </td>" & _
           "    </tr>" & _
           "    <!-- Pie de página -->" & _
           "    <tr>" & _
           "      <td style='background-color: #F1F5F9; padding: 18px 30px; text-align: center; font-size: 12px; color: #64748B; border-top: 1px solid #E2E8F0;'>" & _
           "        Este es un recordatorio automático generado desde el gestor de experimentos del laboratorio.<br>" & _
           "        Si tienes dudas o algún impedimento con los requerimientos, comunícate de inmediato con el responsable." & _
           "      </td>" & _
           "    </tr>" & _
           "  </table>" & _
           "</body>" & _
           "</html>"
           
    GenerarPlantillaHTML = html
End Function

' ==============================================================================
' SUB OPCIONAL: ProbarEnvioFilaSeleccionada
' ==============================================================================
' Permite al usuario seleccionar cualquier celda de una fila y enviar un correo
' de prueba solo para esa fila, sin importar la fecha, mostrando el correo
' en pantalla (Display) antes de enviar. Ideal para validar formato.
' ==============================================================================
Sub ProbarEnvioFilaSeleccionada()
    Dim ws As Worksheet
    Dim filaActual As Long
    Dim tituloExp As String
    Dim responsableExp As String
    Dim fechaExp As Variant
    Dim horaExp As String
    Dim ubicacionExp As String
    Dim checklistTexto As String
    Dim correosDestino As String
    Dim olApp As Object
    Dim olMail As Object
    
    Set ws = ActiveSheet
    filaActual = ActiveCell.Row
    
    If filaActual < FILA_INICIAL_DATOS Then
        MsgBox "Por favor selecciona una celda dentro de la fila del experimento que deseas probar.", vbExclamation, "Fila no válida"
        Exit Sub
    End If
    
    tituloExp = Trim(CStr(ws.Cells(filaActual, COL_TITULO).Value))
    fechaExp = ws.Cells(filaActual, COL_FECHA).Value
    correosDestino = Trim(CStr(ws.Cells(filaActual, COL_CORREOS).Value))
    responsableExp = Trim(CStr(ws.Cells(filaActual, COL_LIDER).Value))
    horaExp = Trim(CStr(ws.Cells(filaActual, COL_HORA).Value))
    ubicacionExp = Trim(CStr(ws.Cells(filaActual, COL_UBICACION).Value))
    checklistTexto = Trim(CStr(ws.Cells(filaActual, COL_CHECKLIST).Value))
    
    If tituloExp = "" Then
        MsgBox "La fila seleccionada no tiene un título de experimento válido.", vbExclamation, "Datos incompletos"
        Exit Sub
    End If
    
    If Not IsDate(fechaExp) Then fechaExp = Date + 1
    If horaExp = "" Then horaExp = "10:00 AM"
    If ubicacionExp = "" Then ubicacionExp = "Laboratorio de Pruebas"
    
    On Error Resume Next
    Set olApp = GetObject(, "Outlook.Application")
    If olApp Is Nothing Then Set olApp = CreateObject("Outlook.Application")
    On Error GoTo 0
    
    If olApp Is Nothing Then
        MsgBox "No se pudo iniciar Outlook.", vbCritical
        Exit Sub
    End If
    
    Set olMail = olApp.CreateItem(0)
    With olMail
        .To = correosDestino
        If .To = "" Then .To = olApp.Session.CurrentUser.Address ' Si está vacío, se envía a uno mismo
        .Subject = "🔔 [PRUEBA] Recordatorio: Experimento '" & tituloExp & "' - Checklist y Requerimientos"
        .HTMLBody = GenerarPlantillaHTML(tituloExp, responsableExp, CDate(fechaExp), horaExp, ubicacionExp, checklistTexto)
        .Display ' Abre la ventana del correo para previsualizarlo
    End With
End Sub

' ==============================================================================
' PROCEDIMIENTO MULTIHERRAMIENTA WORD: EnviarUsandoPlantillaWord
' ==============================================================================
' Variante avanzada multiherramienta: Usa un documento de Word (.docx) como
' plantilla oficial para el correo. Reemplaza marcadores [TITULO], [FECHA], etc.
' y exporta el contenido con el formato exacto de Word al correo de Outlook.
' ==============================================================================
Sub EnviarUsandoPlantillaWord()
    Dim ws As Worksheet
    Dim rutaPlantillaWord As String
    Dim dlg As FileDialog
    Dim wrdApp As Object
    Dim wrdDoc As Object
    Dim olApp As Object
    Dim olMail As Object
    Dim i As Long, ultimaFila As Long
    Dim diasDiferencia As Long
    Dim fechaExp As Variant
    Dim estadoEnvio As String
    Dim tituloExp As String
    Dim totalEnviados As Long
    
    Set ws = ThisWorkbook.Sheets(NOMBRE_HOJA_EXPERIMENTOS)
    ultimaFila = ws.Cells(ws.Rows.Count, COL_TITULO).End(-4162).Row
    
    ' 1. Seleccionar archivo de plantilla Word
    Set dlg = Application.FileDialog(msoFileDialogFilePicker)
    With dlg
        .Title = "Selecciona la plantilla de Word para el Checklist (.docx)"
        .Filters.Clear
        .Filters.Add "Documentos de Word", "*.docx;*.docm;*.doc"
        .AllowMultiSelect = False
        If .Show = -1 Then
            rutaPlantillaWord = .SelectedItems(1)
        Else
            MsgBox "Operación cancelada por el usuario.", vbInformation
            Exit Sub
        End If
    End With
    
    ' 2. Iniciar aplicaciones Word y Outlook
    Set wrdApp = CreateObject("Word.Application")
    wrdApp.Visible = False
    
    Set olApp = CreateObject("Outlook.Application")
    totalEnviados = 0
    
    For i = FILA_INICIAL_DATOS To ultimaFila
        tituloExp = Trim(CStr(ws.Cells(i, COL_TITULO).Value))
        fechaExp = ws.Cells(i, COL_FECHA).Value
        estadoEnvio = UCase(Trim(CStr(ws.Cells(i, COL_ESTADO_ENVIO).Value)))
        
        If tituloExp <> "" And IsDate(fechaExp) Then
            diasDiferencia = DateDiff("d", Date, CDate(fechaExp))
            
            If diasDiferencia = 1 And InStr(estadoEnvio, "ENVIADO") = 0 Then
                ' Abrir una copia limpia de la plantilla Word en modo solo lectura
                Set wrdDoc = wrdApp.Documents.Open(rutaPlantillaWord, ReadOnly:=True)
                
                ' Reemplazar etiquetas en el documento Word
                With wrdDoc.Content.Find
                    .Execute FindText:="{{TITULO}}", ReplaceWith:=tituloExp, Replace:=2 ' 2 = wdReplaceAll
                    .Execute FindText:="{{RESPONSABLE}}", ReplaceWith:=CStr(ws.Cells(i, COL_LIDER).Value), Replace:=2
                    .Execute FindText:="{{FECHA}}", ReplaceWith:=Format(fechaExp, "dd/mm/yyyy"), Replace:=2
                    .Execute FindText:="{{HORA}}", ReplaceWith:=CStr(ws.Cells(i, COL_HORA).Value), Replace:=2
                    .Execute FindText:="{{UBICACION}}", ReplaceWith:=CStr(ws.Cells(i, COL_UBICACION).Value), Replace:=2
                    .Execute FindText:="{{CHECKLIST}}", ReplaceWith:=CStr(ws.Cells(i, COL_CHECKLIST).Value), Replace:=2
                End With
                
                ' Copiar el contenido formateado desde Word al portapapeles
                wrdDoc.Content.Copy
                
                ' Crear correo en Outlook usando el editor de Word
                Set olMail = olApp.CreateItem(0)
                With olMail
                    .To = Trim(CStr(ws.Cells(i, COL_CORREOS).Value))
                    .Subject = "🔔 RECORDATORIO [1 DÍA]: " & tituloExp & " - Checklist de Preparación"
                    .Display ' Necesario para acceder al WordEditor de Outlook
                    .GetInspector.WordEditor.Range(0, 0).Paste
                    .Send
                End With
                
                wrdDoc.Close SaveChanges:=False
                
                ws.Cells(i, COL_ESTADO_ENVIO).Value = "ENVIADO (WORD)"
                ws.Cells(i, COL_FECHA_ENVIO).Value = Now
                totalEnviados = totalEnviados + 1
            End If
        End If
    Next i
    
    wrdApp.Quit
    Set wrdDoc = Nothing
    Set wrdApp = Nothing
    
    MsgBox "Proceso completado. Recordatorios enviados usando plantilla Word: " & totalEnviados, vbInformation, "Listo"
End Sub
