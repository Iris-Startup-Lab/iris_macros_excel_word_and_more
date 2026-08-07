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
' Depende de la función GENERAR_UUID() (módulo macro_generadora_uuid.vb).

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

' Pública porque M_Blindaje también guarda ahí el registro de posiciones de los
' botones (columnas D:H).
Public Function HojaMeta() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Meta")
    On Error GoTo 0

    If ws Is Nothing Then
        ' La estructura del libro está protegida (para que nadie cree hojas de
        ' cotización). Aquí sí necesitamos crear una hoja de SISTEMA: avisamos
        ' con la bandera para que Workbook_NewSheet no la borre.
        gCreandoHojaSistema = True
        DesprotegerEstructura

        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Meta"
        ws.Range("A1").Value = "PROYECTO"
        ws.Range("A2").Value = "UUID"
        ws.Range("A3").Value = "ACTUALIZADO"
        ws.Range("D1").Value = ""        ' D:H = registro de posiciones de botones
    End If

    On Error Resume Next
    ws.Visible = xlSheetVeryHidden   ' no aparece ni en el menú de "Mostrar hoja"
    On Error GoTo 0

    If gCreandoHojaSistema Then
        ProtegerEstructura
        gCreandoHojaSistema = False
    End If

    Set HojaMeta = ws
End Function

' UUID que el libro tiene guardado para el proyecto actual, sin generar nada.
' Devuelve "" si todavía no hay ninguno.
Public Function UUIDGuardado() As String
    Dim meta As Worksheet
    Set meta = HojaMeta()
    If Not meta Is Nothing Then UUIDGuardado = Trim$(CStr(meta.Range("B2").Value))
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

    ' Meta es la fuente de la verdad: si K4 está vacío (lo borraron) o alguien lo
    ' escribió a mano, se REIMPRIME el mismo UUID. No se inventa uno nuevo.
    ws.Range("K4").NumberFormat = "@"   ' texto, para que Excel no lo deforme
    If CStr(ws.Range("K4").Value) <> uuidGuardado Then
        ws.Range("K4").Value = uuidGuardado
    End If
End Sub

' Vigilante del ID Cotización. Se llama desde Worksheet_Change de la hoja
' Cotización y cubre los dos casos:
'   - Cambió C4 (nombre del proyecto)  -> se asigna el UUID que toca.
'   - Se borró o se editó K4 a mano    -> se REIMPRIME el mismo UUID y se avisa.
' Es reentrante-seguro (Static ocupado) y siempre deja EnableEvents en True.
Public Sub VigilarIDCotizacion(ByVal Target As Range)
    Static ocupado As Boolean
    Dim ws As Worksheet
    Dim antes As String, despues As String
    Dim tocaronID As Boolean

    If ocupado Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Target.Worksheet.Name <> ws.Name Then Exit Sub
    If Intersect(Target, ws.Range("C4,K4")) Is Nothing Then Exit Sub

    tocaronID = Not (Intersect(Target, ws.Range("K4")) Is Nothing)

    ocupado = True
    On Error GoTo salir
    Application.EnableEvents = False

    antes = Trim$(CStr(ws.Range("K4").Value))
    AsegurarUUIDCotizacion
    despues = Trim$(CStr(ws.Range("K4").Value))

    Application.EnableEvents = True
    ocupado = False

    ' Solo avisamos cuando el usuario tocó K4 y tuvimos que reponerlo.
    If tocaronID And StrComp(antes, despues, vbTextCompare) <> 0 Then
        If Len(despues) > 0 Then
            MsgBox "El ID Cotización no se edita ni se borra a mano: " & _
                   "se repuso automáticamente." & vbCrLf & vbCrLf & _
                   "ID Cotización: " & despues, vbExclamation, "ID Cotización"
        Else
            MsgBox "No hay ID Cotización porque falta el nombre del proyecto (C4)." & _
                   vbCrLf & "Escribe el proyecto y el ID se generará solo.", _
                   vbExclamation, "ID Cotización"
        End If
    End If
    Exit Sub

salir:
    Application.EnableEvents = True
    ocupado = False
End Sub

' Botón/red de seguridad: repone el ID Cotización si falta, sin preguntar nada.
' Se llama al abrir el libro, al guardar y antes de generar los PDFs.
Public Sub RestaurarIDSiFalta()
    Dim ws As Worksheet
    Set ws = HojaCotizacion()
    If ws Is Nothing Then Exit Sub
    If Len(Trim$(CStr(ws.Range("K4").Value))) > 0 Then
        If StrComp(Trim$(CStr(ws.Range("K4").Value)), UUIDGuardado(), vbTextCompare) = 0 Then Exit Sub
    End If
    On Error Resume Next
    Application.EnableEvents = False
    AsegurarUUIDCotizacion
    Application.EnableEvents = True
    On Error GoTo 0
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
