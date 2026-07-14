Attribute VB_Name = "M_UUID"
' Autor: Fernando Dorantes Nieto
'
' MÓDULO: M_UUID
' Propósito: asignar automáticamente un UUID (ID Cotización, celda K4) que:
'   - Es único por cotización.
'   - Queda ATADO al nombre del proyecto (celda C4). Solo se regenera si C4 cambia.
'   - Se MANTIENE aunque cambie la versión (N4) o se reabra el archivo.
'   - Es la misma cadena que luego se usa en el nombre del archivo PDF, de modo
'     que archivo e "ID Cotización" siempre coinciden.
'
' El par (proyecto, uuid) se guarda en una hoja MUY oculta llamada "Meta",
' que la propia macro crea la primera vez. Así viaja dentro del libro.
'
' Depende de la función GENERAR_UUID() (módulo M_GeneradorUUID).

Option Explicit

' --- Configuración compartida por los demás módulos ---
Public Const COL_MARCA As String = "O"       ' columna oculta con la marca de detalle
Public Const MARCA_DETALLE As String = "D"    ' valor que marca una fila como "detalle colapsable"
Public Const FILA_INICIO As Long = 7          ' primera fila de bloques de categorías

' --- Helpers de localización de hojas (robustos ante espacios/acentos) ---
Public Function HojaQueContiene(ByVal token As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If InStr(1, UCase$(Trim$(ws.Name)), UCase$(token)) > 0 Then
            Set HojaQueContiene = ws
            Exit Function
        End If
    Next ws
End Function

Public Function HojaCotizacion() As Worksheet
    Set HojaCotizacion = HojaQueContiene("COTIZACI")
End Function

Private Function HojaMeta() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Meta")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Meta"
        ws.Range("A1").Value = "PROYECTO"
        ws.Range("A2").Value = "UUID"
        ws.Range("A3").Value = "ACTUALIZADO"
    End If

    ws.Visible = xlSheetVeryHidden   ' no aparece ni en el menú de "Mostrar hoja"
    Set HojaMeta = ws
End Function

' --- Núcleo: garantiza que K4 tenga el UUID correcto para el proyecto actual ---
Public Sub AsegurarUUIDCotizacion()
    Dim ws As Worksheet, meta As Worksheet
    Dim proyecto As String, proyGuardado As String, uuidGuardado As String

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub

    proyecto = Trim$(CStr(ws.Range("C4").Value))
    ' Sin nombre de proyecto todavía no generamos nada (evita UUIDs "huérfanos").
    If Len(proyecto) = 0 Then Exit Sub

    Set meta = HojaMeta()
    proyGuardado = Trim$(CStr(meta.Range("B1").Value))
    uuidGuardado = Trim$(CStr(meta.Range("B2").Value))

    ' Regeneramos solo si: no hay UUID, o el proyecto cambió.
    If Len(uuidGuardado) = 0 Or StrComp(proyGuardado, proyecto, vbTextCompare) <> 0 Then
        uuidGuardado = GENERAR_UUID()
        meta.Range("B1").Value = proyecto
        meta.Range("B2").Value = uuidGuardado
        meta.Range("B3").Value = Now
    End If

    ws.Range("K4").NumberFormat = "@"   ' texto, para que Excel no lo deforme
    If CStr(ws.Range("K4").Value) <> uuidGuardado Then
        ws.Range("K4").Value = uuidGuardado
    End If
End Sub

' Devuelve el UUID actual (asegurándolo primero). Útil para otros módulos.
Public Function UUIDActual() As String
    AsegurarUUIDCotizacion
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If Not ws Is Nothing Then UUIDActual = Trim$(CStr(ws.Range("K4").Value))
End Function

' Botón manual opcional: fuerza revisión/asignación del UUID.
Public Sub AsignarUUID_Click()
    AsegurarUUIDCotizacion
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then
        MsgBox "No se encontró la hoja de Cotización.", vbCritical
    Else
        MsgBox "ID Cotización (UUID): " & vbCrLf & ws.Range("K4").Value, vbInformation
    End If
End Sub
